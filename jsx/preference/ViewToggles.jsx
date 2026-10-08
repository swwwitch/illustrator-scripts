#target illustrator
#targetengine "ViewTogglesEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アプリケーションフレーム・コントロールパネル・ガイド・グリッド・スナップなどの表示とオン・オフを、ダイアログボックスのチェックボックスでまとめて切り替えます。
今の状態を読んでから開くので、メニューの反転と違って意図しない切り替えが起きません。

### 注意

メニューの状態を読むには補助アプリ /Applications/SetAiMenuState.app（ai-scripts の helpers/SetAiMenuState.applescript）とアクセシビリティの許可が要ります。
補助アプリが無いときは、環境設定で読める項目だけを切り替えられます。

### Overview

Switches the application frame, Control panel, guides, grid, snapping and more on or off with checkboxes in one dialog.
The current state is read first, so nothing is flipped by accident as with a plain menu toggle.

### Notes

Reading menu states needs the helper app /Applications/SetAiMenuState.app (helpers/SetAiMenuState.applescript in ai-scripts) with Accessibility permission.
Without it, only the items readable from preferences can be switched.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ViewToggles";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-09";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ツールバーをオンにするときに表示するもの: "quick"（はじめに）/ "basic"（基本）/ "advanced"（詳細）
       Toolbar shown when turning the toolbar on */
    var TOOLBAR_TO_SHOW = "basic";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var LEFT_COLUMN_GROUP_COUNT = 3;         /* 左の列に置くカテゴリの数（残りは右の列）/ categories in the left column (the rest go right) */
    var ROW_SPACING           = 8;           /* 行内の要素間隔 / spacing inside a row */
    var PRESET_DROPDOWN_WIDTH = 160;         /* プリセットのドロップダウンの幅 / preset dropdown width */
    var PRESET_ICON_SIZE      = [20, 18];    /* プリセットの保存・削除アイコンの大きさ / preset save and delete icon size */
    var SAVE_ICON_OPACITY     = 0.75;        /* 保存アイコンの濃さ（ほかのアイコンに対する比率）/ save icon opacity relative to the others */
    var PRESET_NAME_CHARS     = 20;          /* プリセット名の入力欄の文字数 / preset name field width */
    var BRIGHTNESS_SWATCH_SIZE    = 30;      /* 明るさの見本1つの大きさ（選択枠を含む）/ size of one brightness swatch, selection ring included */
    var BRIGHTNESS_SWATCH_SPACING = 4;       /* 明るさの見本どうしの間隔 / gap between brightness swatches */
    var BRIGHTNESS_RING_COLOR     = [0.29, 0.56, 0.95, 1]; /* 選んでいる見本の枠の色 / ring color of the selected swatch */
    var BRIGHTNESS_EDGE_COLOR     = [0.85, 0.85, 0.85, 1]; /* 見本の縁の色 / edge color of each swatch */

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

    // アイコンのボタン（再利用パーツ） / Icon buttons (reusable)

    var ICON_BUTTON_SIZE = [24, 22]; /* 既定の大きさ / default size */
    var ICON_BUTTON_UI_DARK = isDarkUI();
    var ICON_BUTTON_COLOR     = ICON_BUTTON_UI_DARK ? [1, 1, 1, 1]    : [0, 0, 0, 0.70]; /* 絵の色 / icon color */
    var ICON_BUTTON_DIM_COLOR = ICON_BUTTON_UI_DARK ? [1, 1, 1, 0.20] : [0, 0, 0, 0.25]; /* 無効時の色 / color when disabled */

    /**
     * onDraw で絵を描くアイコンのボタンを追加する。押すと onClick を呼ぶ（無効の間は押せず、薄く描く）
     * @param {Group} parent - 追加先
     * @param {number[]} iconSize - [幅, 高さ]
     * @param {Function} drawIcon - 絵を描く関数 (iconGraphics, iconWidth, iconHeight, iconColor)
     * @returns {Group} アイコン
     */
    function addIconButton(parent, iconSize, drawIcon) {
        var iconButton = parent.add("group");
        iconButton.preferredSize = iconSize;
        iconButton.minimumSize = iconSize;
        iconButton.maximumSize = iconSize;
        iconButton.onDraw = function () {
            var iconColor = isIconButtonEnabledInTree(iconButton) ? ICON_BUTTON_COLOR : ICON_BUTTON_DIM_COLOR;
            drawIcon(iconButton.graphics, iconSize[0], iconSize[1], iconColor);
        };
        iconButton.addEventListener("mousedown", function () {
            if (!isIconButtonEnabledInTree(iconButton)) return;
            if (typeof iconButton.onClick === "function") iconButton.onClick();
        });
        return iconButton;
    }

    /**
     * アイコンのボタンの有効／無効を切り替えて描き直す（変わらないときは描き直さない）
     * @param {Group} iconButton - addIconButton() で作ったアイコン
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setIconButtonEnabled(iconButton, isEnabled) {
        if (iconButton.enabled === isEnabled) return;
        iconButton.enabled = isEnabled;
        /* group には notify() が無いため、隠して再表示して描き直させる / groups have no notify(), so hide and show to repaint */
        iconButton.hide();
        iconButton.show();
    }

    /**
     * コントロールと親がすべて有効かを判定する（親の無効化は子の enabled に出ないため、親もたどる）
     * @param {Object} control - 判定するコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isIconButtonEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * 絵の座標の長方形の並びを、縦横比を保って中央に置いて塗る（ScriptUI は多角形を塗れないため、斜めも長方形の並びで描く）
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} designSize - 絵の [幅, 高さ]
     * @param {number[][]} designRects - 長方形 [左, 上, 右, 下] の並び
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function fillIconButtonRects(iconGraphics, iconWidth, iconHeight, designSize, designRects, iconColor) {
        var iconScale = Math.min(iconWidth / designSize[0], iconHeight / designSize[1]);
        var originX = (iconWidth - designSize[0] * iconScale) / 2;
        var originY = (iconHeight - designSize[1] * iconScale) / 2;
        iconGraphics.newPath();
        for (var i = 0; i < designRects.length; i++) {
            var designRect = designRects[i];
            if (designRect[2] <= designRect[0] || designRect[3] <= designRect[1]) continue;
            iconGraphics.rectPath(originX + designRect[0] * iconScale, originY + designRect[1] * iconScale,
                (designRect[2] - designRect[0]) * iconScale, (designRect[3] - designRect[1]) * iconScale);
        }
        iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, iconColor));
    }

    /**
     * 保存アイコン（トレイに下向きの矢印）を描く。1223×993 の絵。
     * 矢じりとトレイの V 字の切り欠きは細い長方形を並べ、トレイの四角い穴は塗らずに残す
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawSaveIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        var designWidth = 1223;
        var trayHole = [78, 688, 230, 840]; /* トレイの四角い穴 [左, 上, 右, 下] / the tray's square hole */
        var sliceCount = 16;
        var designRects = [[535, 0, 688, 383]]; /* 矢印の軸 / arrow shaft */
        var k;

        /** 穴と重なる部分を除いて長方形を足す / add a rectangle minus the tray hole */
        function addTrayRect(left, top, right, bottom) {
            if (right <= trayHole[0] || left >= trayHole[2] || bottom <= trayHole[1] || top >= trayHole[3]) {
                designRects.push([left, top, right, bottom]);
                return;
            }
            designRects.push([left, top, right, trayHole[1]], [left, trayHole[3], right, bottom],
                [left, Math.max(top, trayHole[1]), trayHole[0], Math.min(bottom, trayHole[3])],
                [trayHole[2], Math.max(top, trayHole[1]), right, Math.min(bottom, trayHole[3])]);
        }

        /* 下向きの矢じり（上の付け根から先端へ細っていく）/ the downward head */
        for (k = 0; k < sliceCount; k++) {
            var headHalf = 251 * (1 - (k + 0.5) / sliceCount);
            var headTop = 383 + (703 - 383) * k / sliceCount;
            designRects.push([612 - headHalf, headTop, 612 + headHalf, headTop + (703 - 383) / sliceCount]);
        }
        /* トレイ：上辺に V 字の切り欠き（幅 458 から下へ細り、856 で閉じる）/ tray with a V notch along the top */
        for (k = 0; k < sliceCount; k++) {
            var notchTop = 612 + (856 - 612) * k / sliceCount;
            var notchBottom = notchTop + (856 - 612) / sliceCount;
            var notchHalf = 229 * (1 - (k + 0.5) / sliceCount);
            addTrayRect(0, notchTop, 612 - notchHalf, notchBottom);
            addTrayRect(612 + notchHalf, notchTop, designWidth, notchBottom);
        }
        addTrayRect(0, 856, designWidth, 993);
        fillIconButtonRects(iconGraphics, iconWidth, iconHeight, [designWidth, 993], designRects, iconColor);
    }

    /**
     * 削除アイコン（ゴミ箱）を描く。756×825 の絵
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawTrashIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        fillIconButtonRects(iconGraphics, iconWidth, iconHeight, [756, 825], [
            [206, 0, 550, 70], [206, 70, 275, 137], [481, 70, 550, 137],     /* 取っ手 / handle */
            [0, 137, 756, 207],                                               /* ふた / lid */
            [69, 207, 138, 825], [618, 207, 688, 825], [138, 756, 618, 825], /* 本体 / body */
            [206, 275, 275, 687], [343, 275, 413, 687], [481, 275, 550, 687] /* 縦の線 / ribs */
        ], iconColor);
    }

    // アイコンのボタン（再利用パーツ）ここまで / End of the reusable icon buttons

    // =========================================
    // 設定の保存 / Settings store
    // =========================================

    // 設定の保存（再利用パーツ） / Settings store (reusable)

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
            return textFile.read().replace(/^\uFEFF/, "");
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
        var trimmedText = legacyText.replace(/^\uFEFF/, "").replace(/^\s+|\s+$/g, "");
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

    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store

    /* プリセットは再起動しても残す / Presets survive a restart */
    var presetSettingsStore = createSettingsStore(SCRIPT_NAME + "Presets", "persistent");

    // =========================================
    // メニューの状態 / Menu states
    // =========================================

    // メニューの ✓ を読む・揃える（再利用パーツ） / Read and set menu checkmarks (reusable)

    /**
     * 補助アプリ /Applications/SetAiMenuState.app に依頼を書いて起動する
     * @param {Array} requests - [メニューの道筋, "get"|"on"|"off"|"toggle"] の配列
     * @returns {boolean} 補助アプリを起動できたら true。無い・起動できない・macOS 以外のときは false
     */
    function launchAiMenuStateHelper(requests) {
        /* 定数は巻き上げで未定義にならないよう関数内に置く / Kept local so hoisting never leaves them undefined */
        var helperAppPath = "/Applications/SetAiMenuState.app";
        var requestFilePath = "/tmp/set_ai_menu_state_request.txt";

        if ($.os.indexOf("Mac") === -1) return false;
        /* .app は実体がディレクトリなので Folder でも確かめる / An .app is a directory, so check it as a Folder too */
        if (!new Folder(helperAppPath).exists && !new File(helperAppPath).exists) return false;

        var lines = [];
        for (var i = 0; i < requests.length; i++) {
            lines.push(requests[i][0] + "\t" + requests[i][1]);
        }

        var requestFile = new File(requestFilePath);
        var written = false;
        try {
            requestFile.encoding = "UTF-8";
            requestFile.lineFeed = "Unix";
            if (requestFile.open("w")) {
                written = requestFile.write(lines.join("\n"));
            }
        } catch (e) {
        } finally {
            try { requestFile.close(); } catch (closeError) {}
        }
        return written && new File(helperAppPath).execute();
    }

    /**
     * メニュー項目の ✓ を読む。
     * スクリプトの実行中はメニューを読めないため、補助アプリが読み終えるまで小さなダイアログボックスを出して待つ
     * @param {string[]} menuPaths - 「表示>スマートガイド」の形のメニューの道筋
     * @returns {Object|null} 道筋をキーに "on" / "off" / "notfound" を持つオブジェクト。読めなかったら null
     */
    function getAiMenuState(menuPaths) {
        /* 補助アプリはこの名前の UI element を探してボタンを押す / The helper looks up a UI element with this name and presses its button */
        var waitDialogName = "SetAiMenuState";
        var resultFile = new File("/tmp/set_ai_menu_state_result.txt");
        var isJapanese = $.locale.indexOf("ja") === 0;

        /* 前回の結果を読まないように消しておく / Remove the previous result so it is never read by mistake */
        if (resultFile.exists) resultFile.remove();

        var requests = [];
        for (var i = 0; i < menuPaths.length; i++) requests.push([menuPaths[i], "get"]);
        if (!launchAiMenuStateHelper(requests)) return null;

        var waitDialog = new Window("dialog", waitDialogName);
        waitDialog.add("statictext", undefined, isJapanese ? "メニューの状態を確認中..." : "Checking menu states...");
        /* 補助アプリが動かないときは手で閉じる / Closed by hand when the helper cannot run */
        var cancelButton = waitDialog.add("button", undefined, isJapanese ? "キャンセル" : "Cancel");
        cancelButton.onClick = function () { waitDialog.close(); };
        waitDialog.show();

        if (!resultFile.exists) return null;
        var resultText = "";
        try {
            resultFile.encoding = "UTF-8";
            if (resultFile.open("r")) resultText = resultFile.read();
        } catch (e) {
        } finally {
            try { resultFile.close(); } catch (closeError) {}
        }

        /* 1行は「道筋<TAB>変更前<TAB>変更後」。error 行があれば失敗 / Each line is "path<TAB>before<TAB>after"; an error line means failure */
        var states = {};
        var resultLines = resultText.split("\n");
        for (var j = 0; j < resultLines.length; j++) {
            var fields = resultLines[j].split("\t");
            if (fields[0] === "error") return null;
            if (fields.length >= 2) states[fields[0]] = fields[1];
        }
        return states;
    }

    /**
     * メニュー項目の ✓ をオン・オフに揃えるよう、補助アプリに依頼する。
     * 切り替えはこのスクリプトが終わってから行われる
     * @param {Array} requests - [メニューの道筋, "on"|"off"|"toggle"] の配列。道筋は「表示>スマートガイド」の形
     * @returns {boolean} 補助アプリを起動できたら true。無い・起動できない・macOS 以外のときは false
     */
    function setAiMenuState(requests) {
        return launchAiMenuStateHelper(requests);
    }

    // メニューの ✓ を読む・揃える（再利用パーツ）ここまで / End of the reusable menu checkmark helpers

    // =========================================
    // 項目の定義 / Item definitions
    // =========================================

    /* ツールバーの種類ごとのメニュー名。切り替えは補助アプリがメニューをクリックする
       （コマンド ID は反転ではなく別のツールバーに切り替わることがあった。2026-10-09 実測）
       Menu names per toolbar. The helper clicks the menu (a command ID could switch to another toolbar instead of toggling) */
    var TOOLBAR_KINDS = [
        { kind: "quick",    menuName: { ja: "はじめに", en: "Quick" } },
        { kind: "basic",    menuName: { ja: "基本",     en: "Basic" } },
        { kind: "advanced", menuName: { ja: "詳細",     en: "Advanced" } }
    ];

    /* ［ユーザーインターフェイス］の明るさの4段階。value は uiBrightness の値、prefButton は環境設定のボタンの名前、
       fill は見本の色。0.5 と 0.51 は別の段階なので、いちばん近い値で判定する
       The four Brightness levels: value is uiBrightness, prefButton the button name in Preferences, fill the swatch color.
       0.5 and 0.51 are different levels, so the nearest value wins */
    var BRIGHTNESS_LEVELS = [
        { key: "dark",        value: 0.0,  prefButton: "暗",         fill: [0.20, 0.20, 0.20, 1] },
        { key: "mediumDark",  value: 0.5,  prefButton: "やや暗め",   fill: [0.33, 0.33, 0.33, 1] },
        { key: "mediumLight", value: 0.51, prefButton: "やや明るめ", fill: [0.72, 0.72, 0.72, 1] },
        { key: "light",       value: 1.0,  prefButton: "明",         fill: [0.94, 0.94, 0.94, 1] }
    ];

    /*
     * 項目ごとに、状態の読み方と切り替え方を決める。
     *   pref    : 環境設定キーで読む（メニューの切り替えにすぐ追従する。2026-10-09 実測）
     *   menu    : 補助アプリでメニューの ✓ や項目名を読む。「オフの名前|オンの名前」は名前が入れ替わる項目
     *   command : executeMenuCommand() の ID。無い項目は補助アプリがメニューをクリックする（スクリプトの終了後）
     *   prefType : "integer" なら整数として読む（真偽として読むとオフでも true が返る）。省略時は真偽
     *   writePref : 切り替えも環境設定キーに書いて行う（書けばすぐ効く項目）
     * 英語版のメニュー名は未検証
     * How each item is read and switched:
     *   pref    : read from a preference key (follows the menu immediately; measured 2026-10-09)
     *   menu    : read by the helper from the checkmark or the label; "off name|on name" marks a label that flips
     *   command : the executeMenuCommand() ID; without one the helper clicks the menu after the script ends
     *   prefType : "integer" reads the key as an integer (read as a boolean it is true even when off); boolean otherwise
     *   writePref : switched by writing the preference key too (for keys that take effect at once)
     * The English menu names are unverified
     */
    var VIEW_TOGGLE_GROUPS = [
        {
            labelKey: "basicUI",
            hasBrightness: true, /* 先頭に明るさの見本を置く / brightness swatches come first */
            items: [
                { key: "appFrame",    menu: { ja: "ウィンドウ>アプリケーションフレーム", en: "Window>Application Frame" }, command: "appframe" },
                { key: "appBar",      menu: { ja: "ウィンドウ>アプリケーションバー", en: "Window>Application Bar" }, command: "applicationbar" },
                { key: "taskBar",     pref: "ContextualTaskBarEnabled", prefType: "integer", menu: { ja: "ウィンドウ>コンテキストタスクバー", en: "Window>Contextual Task Bar" } },
                { key: "controlBar",  menu: { ja: "ウィンドウ>コントロール", en: "Window>Control" }, command: "drover control palette plugin" },
                { key: "toolbar",     toolbar: true },
                { key: "helpBar",     pref: "showHelpBar", prefType: "integer", menu: { ja: "ウィンドウ>ヘルプバー", en: "Window>Help Bar" } },
                { key: "genMenu",     menu: { ja: "ウィンドウ>生成メニューを表示|生成メニューを非表示", en: "Window>Show Generative Menu|Hide Generative Menu" } }
            ]
        },
        {
            labelKey: "objects",
            items: [
                { key: "cornerWidget", pref: "liveCorners/showWidget", command: "Live Corner Annotator" },
                { key: "edges",        menu: { ja: "表示>境界線を表示|境界線を隠す", en: "View>Show Edges|Hide Edges" }, command: "edge" },
                { key: "boundingBox",  pref: "showBoundingBox", command: "AI Bounding Box Toggle" }
            ]
        },
        {
            labelKey: "artboard",
            items: [
                { key: "artboards",     menu: { ja: "表示>アートボードを表示|アートボードを隠す", en: "View>Show Artboards|Hide Artboards" }, command: "artboard" },
                { key: "artboardNames", pref: "showArtboardLabelOnCanvas", writePref: true },
                /* 設定キーを書いてもカンバスのボタンは変わらないので、補助アプリが環境設定のチェックボックスを操作する
                   Writing the key does not update the canvas buttons, so the helper operates the checkbox in the preferences dialog */
                { key: "genAIButton",   pref: "enablePrintBleedWidget", menu: { ja: "環境設定>一般...>裁ち落とし部分に「裁ち落としを印刷」生成 AI ボタンを表示", en: "" } }
            ]
        },
        {
            labelKey: "guides",
            items: [
                { key: "guides",      menu: { ja: "表示>ガイド>ガイドを表示|ガイドを隠す", en: "View>Guides>Show Guides|Hide Guides" }, command: "showguide" },
                { key: "smartGuides", pref: "smartGuides/isEnabled", command: "Snapomatic on-off menu item" },
                { key: "lockGuides",  menu: { ja: "表示>ガイド>ガイドをロック|ガイドをロック解除", en: "View>Guides>Lock Guides|Unlock Guides" }, command: "lockguide" },
                /* ガイドへの吸着は［ポイントにスナップ］が受け持つ / Snapping to guides is handled by Snap to Point */
                { key: "snapPoint",   pref: "snapToPoint", command: "snappoint" }
            ]
        },
        {
            labelKey: "grid",
            items: [
                { key: "grid",     menu: { ja: "表示>グリッドを表示|グリッドを隠す", en: "View>Show Grid|Hide Grid" }, command: "showgrid" },
                { key: "snapGrid", menu: { ja: "表示>グリッドにスナップ", en: "View>Snap to Grid" }, command: "snapgrid" }
            ]
        },
        {
            labelKey: "otherSnap",
            items: [
                { key: "snapPixel", menu: { ja: "表示>ピクセルにスナップ", en: "View>Snap to Pixel" }, command: "pixelconstraints" },
                { key: "snapGlyph", pref: "snapToGlyph", command: "glyphSnapping" }
            ]
        },
        {
            /* JavaScript で設定キーを書いても再起動まで効かないので、補助アプリが環境設定のチェックボックスを操作する
               Keys written from JavaScript only take effect after a restart, so the helper operates the checkboxes in the preferences dialog */
            labelKey: "saveExport",
            items: [
                { key: "backgroundExport", pref: "enableBackgroundExport", menu: { ja: "環境設定>ファイル管理...>バックグラウンドで書き出し", en: "" } },
                { key: "backgroundSave",   pref: "enableBackgroundSave", menu: { ja: "環境設定>ファイル管理...>バックグラウンドで保存", en: "" } },
                { key: "recoveryAutosave", pref: "CrashRecovery/AutomaticallySave", menu: { ja: "環境設定>ファイル管理...>復帰データを次の間隔で自動保存", en: "" } }
            ]
        }
    ];

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
            title: { ja: "表示の切り替え", en: "View Toggles" },
            presetSave: { ja: "プリセットを保存", en: "Save Preset" }
        },
        dropdown: {
            presetPlaceholder: { ja: "---", en: "---" }
        },
        fieldLabel: {
            brightness: { ja: "明るさ", en: "Brightness" },
            preset: { ja: "プリセット", en: "Preset" },
            presetName: { ja: "プリセット名", en: "Preset name" }
        },
        panel: {
            basicUI:       { ja: "基本UI", en: "Basic UI" },
            objects:       { ja: "オブジェクト", en: "Objects" },
            artboard:      { ja: "アートボード", en: "Artboards" },
            guides:        { ja: "ガイド", en: "Guides" },
            grid:          { ja: "グリッド", en: "Grid" },
            otherSnap:     { ja: "その他のスナップ", en: "Other Snapping" },
            saveExport:    { ja: "保存・書き出し", en: "Save and Export" }
        },
        checkbox: {
            appFrame:     { ja: "アプリケーションフレーム", en: "Application Frame" },
            appBar:       { ja: "アプリケーションバー", en: "Application Bar" },
            taskBar:      { ja: "コンテキストタスクバー", en: "Contextual Task Bar" },
            controlBar:   { ja: "コントロールパネル", en: "Control Panel" },
            toolbar:      { ja: "ツールバー", en: "Toolbar" },
            helpBar:      { ja: "ヘルプバー", en: "Help Bar" },
            genMenu:      { ja: "生成メニュー", en: "Generative Menu" },
            cornerWidget: { ja: "コーナーウィジェット", en: "Corner Widget" },
            edges:        { ja: "境界線", en: "Edges" },
            boundingBox:  { ja: "バウンディングボックス", en: "Bounding Box" },
            artboards:     { ja: "境界線の表示", en: "Show Borders" },
            artboardNames: { ja: "アートボード名の表示", en: "Show Artboard Names" },
            genAIButton:   { ja: "生成AIボタンの表示", en: "Show Generative AI Buttons" },
            backgroundExport: { ja: "バックグラウンド書き出し", en: "Export in Background" },
            backgroundSave:   { ja: "バックグラウンド保存", en: "Save in Background" },
            recoveryAutosave: { ja: "復帰データの自動保存", en: "Autosave Recovery Data" },
            guides:       { ja: "ガイドの表示", en: "Show Guides" },
            smartGuides:  { ja: "スマートガイドの表示", en: "Show Smart Guides" },
            lockGuides:   { ja: "ロック", en: "Lock" },
            snapPoint:    { ja: "スナップ", en: "Snap" },
            grid:         { ja: "表示", en: "Show" },
            snapGrid:     { ja: "スナップ", en: "Snap" },
            snapPixel:    { ja: "ピクセル", en: "Pixel" },
            snapGlyph:    { ja: "グリフ", en: "Glyph" }
        },
        tooltip: {
            unavailable: { ja: "状態を読めないため切り替えられません", en: "Unavailable because its state cannot be read" },
            brightness: {
                dark:        { ja: "暗（スクリプトの終了後に環境設定で切り替えます）", en: "Dark (switched in Preferences after the script ends)" },
                mediumDark:  { ja: "やや暗め（スクリプトの終了後に環境設定で切り替えます）", en: "Medium Dark (switched in Preferences after the script ends)" },
                mediumLight: { ja: "やや明るめ（スクリプトの終了後に環境設定で切り替えます）", en: "Medium Light (switched in Preferences after the script ends)" },
                light:       { ja: "明（スクリプトの終了後に環境設定で切り替えます）", en: "Light (switched in Preferences after the script ends)" }
            },
            preset:       { ja: "保存したチェックの組み合わせを読み込みます", en: "Loads a saved set of checkboxes" },
            presetSave:   { ja: "今のチェックの組み合わせに名前を付けて保存します", en: "Saves the current checkboxes under a name" },
            presetDelete: { ja: "選んでいるプリセットを削除します", en: "Deletes the selected preset" },
            presetName:   { ja: "同じ名前のプリセットは上書きします", en: "A preset with the same name is overwritten" },
            taskBar:     { ja: "スクリプトの終了後に切り替わります", en: "Switched after the script ends" },
            helpBar:     { ja: "スクリプトの終了後に切り替わります", en: "Switched after the script ends" },
            genMenu:     { ja: "スクリプトの終了後に切り替わります", en: "Switched after the script ends" },
            toolbar:     { ja: "オンにすると［{toolbarName}］ツールバーを表示します。スクリプトの終了後に切り替わります", en: "Turning it on shows the {toolbarName} toolbar. Switched after the script ends" },
            snapPoint:   { ja: "［表示］→［ポイントにスナップ］です。ガイドとアンカーポイントに吸着します", en: "View > Snap to Point, which snaps to guides and anchor points" },
            snapGrid:    { ja: "［表示］→［グリッドにスナップ］です", en: "View > Snap to Grid" },
            edges:       { ja: "パスの境界線です（［表示］→［境界線を表示］）", en: "Path edges (View > Show Edges)" },
            artboards:   { ja: "アートボードの境界線です（［表示］→［アートボードを表示］）", en: "Artboard borders (View > Show Artboards)" },
            genAIButton: { ja: "裁ち落とし部分の「裁ち落としを印刷」ボタンです。スクリプトの終了後に環境設定を開いて切り替えます", en: "The Print Bleed buttons on the bleed. Switched in Preferences after the script ends" },
            backgroundExport: { ja: "スクリプトの終了後に環境設定を開いて切り替えます", en: "Switched in Preferences after the script ends" },
            backgroundSave:   { ja: "スクリプトの終了後に環境設定を開いて切り替えます", en: "Switched in Preferences after the script ends" },
            recoveryAutosave: { ja: "スクリプトの終了後に環境設定を開いて切り替えます", en: "Switched in Preferences after the script ends" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok:     { ja: "OK", en: "OK" }
        },
        confirm: {
            presetOverwrite: { ja: "「{name}」を上書きしますか？", en: "Overwrite \"{name}\"?" },
            presetDelete:    { ja: "「{name}」を削除しますか？", en: "Delete \"{name}\"?" }
        },
        alert: {
            presetSaveFailed: { ja: "プリセットを保存できませんでした。", en: "Could not save the preset." },
            helperUnavailable: {
                ja: "メニューの状態を読めなかったため、一部の項目は切り替えられません。\n/Applications/SetAiMenuState.app とアクセシビリティの許可を確認してください。",
                en: "Some items are unavailable because the menu states could not be read.\nCheck /Applications/SetAiMenuState.app and its Accessibility permission."
            }
        }
    };

    // =========================================
    // 明るさ / Brightness
    // =========================================

    /**
     * 今の［明るさ］の段階を返す
     * @returns {number} BRIGHTNESS_LEVELS の番号
     */
    function readBrightnessIndex() {
        var brightnessValue = app.preferences.getRealPreference("uiBrightness");
        var nearestIndex = 0;
        for (var i = 1; i < BRIGHTNESS_LEVELS.length; i++) {
            if (Math.abs(BRIGHTNESS_LEVELS[i].value - brightnessValue) < Math.abs(BRIGHTNESS_LEVELS[nearestIndex].value - brightnessValue)) nearestIndex = i;
        }
        return nearestIndex;
    }

    /**
     * 明るさの見本を1つ描く。選んでいれば外側に枠を描く
     * @param {Object} swatchGraphics - 見本の graphics
     * @param {Array} fillColor - 見本の色 [r, g, b, a]
     * @param {boolean} isSelected - 選んでいるなら true
     * @param {boolean} isEnabled - 有効なら true（無効なら薄く描く）
     * @returns {void}
     */
    function drawBrightnessSwatch(swatchGraphics, fillColor, isSelected, isEnabled) {
        var alpha = isEnabled ? 1 : 0.4;
        var swatchSize = BRIGHTNESS_SWATCH_SIZE;
        if (isSelected) {
            /* rectPath の前に newPath() を呼ばないとパスが累積する / Call newPath() before rectPath() or paths accumulate */
            swatchGraphics.newPath();
            swatchGraphics.rectPath(1, 1, swatchSize - 2, swatchSize - 2);
            swatchGraphics.strokePath(swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, BRIGHTNESS_RING_COLOR, 2));
        }
        swatchGraphics.newPath();
        swatchGraphics.rectPath(5, 5, swatchSize - 10, swatchSize - 10);
        swatchGraphics.fillPath(swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR, [fillColor[0], fillColor[1], fillColor[2], alpha]));
        swatchGraphics.strokePath(swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, [BRIGHTNESS_EDGE_COLOR[0], BRIGHTNESS_EDGE_COLOR[1], BRIGHTNESS_EDGE_COLOR[2], alpha], 2));
    }

    /**
     * 「明るさ :」の行に4段階の見本を並べる。クリックで選び、［OK］で環境設定に反映する
     * @param {Panel} parentPanel - 行を足すパネル
     * @param {boolean} isEnabled - 切り替えられるなら true（補助アプリが無い・英語版では false）
     * @returns {{initialIndex: number, selectedIndex: number, isEnabled: boolean, select: Function}} 見本の状態
     */
    function addBrightnessPicker(parentPanel, isEnabled) {
        var brightnessRow = parentPanel.add("group");
        setupRow(brightnessRow, "left", ROW_SPACING);
        brightnessRow.add("statictext", undefined, labelText("fieldLabel.brightness"));
        var swatchRow = brightnessRow.add("group");
        setupRow(swatchRow, "left", BRIGHTNESS_SWATCH_SPACING);

        var currentIndex = readBrightnessIndex();
        var brightnessPicker = { initialIndex: currentIndex, selectedIndex: currentIndex, isEnabled: isEnabled, swatches: [] };

        /* 選んだ見本を変えて、すべて描き直す（group には notify() が無いので hide/show）
           Change the selection and repaint all swatches (groups have no notify(), so hide and show) */
        brightnessPicker.select = function (levelIndex) {
            brightnessPicker.selectedIndex = levelIndex;
            for (var i = 0; i < brightnessPicker.swatches.length; i++) {
                brightnessPicker.swatches[i].hide();
                brightnessPicker.swatches[i].show();
            }
        };

        for (var i = 0; i < BRIGHTNESS_LEVELS.length; i++) {
            brightnessPicker.swatches.push(addBrightnessSwatch(swatchRow, brightnessPicker, i));
        }
        return brightnessPicker;
    }

    /**
     * 明るさの見本を1つ作る
     * @param {Group} swatchRow - 見本を並べる行
     * @param {Object} brightnessPicker - addBrightnessPicker() の状態
     * @param {number} levelIndex - BRIGHTNESS_LEVELS の番号
     * @returns {Group} 見本
     */
    function addBrightnessSwatch(swatchRow, brightnessPicker, levelIndex) {
        var brightnessLevel = BRIGHTNESS_LEVELS[levelIndex];
        var swatchGroup = swatchRow.add("group");
        var swatchSize = [BRIGHTNESS_SWATCH_SIZE, BRIGHTNESS_SWATCH_SIZE];
        swatchGroup.preferredSize = swatchSize;
        swatchGroup.minimumSize = swatchSize;
        swatchGroup.maximumSize = swatchSize;
        swatchGroup.helpTip = brightnessPicker.isEnabled ? getLabel("tooltip.brightness." + brightnessLevel.key) : getLabel("tooltip.unavailable");
        swatchGroup.onDraw = function () {
            drawBrightnessSwatch(swatchGroup.graphics, brightnessLevel.fill, brightnessPicker.selectedIndex === levelIndex, brightnessPicker.isEnabled);
        };
        swatchGroup.addEventListener("mousedown", function () {
            if (brightnessPicker.isEnabled) brightnessPicker.select(levelIndex);
        });
        return swatchGroup;
    }

    // =========================================
    // プリセット / Presets
    // =========================================

    /**
     * 保存したプリセットを返す
     * @returns {Object} プリセット名をキーにしたチェックの組み合わせ（書き換えても保存されない）
     */
    function loadPresetMap() {
        return presetSettingsStore.load({});
    }

    /**
     * 保存したプリセットの名前を、名前順で返す
     * @returns {string[]} プリセット名
     */
    function getPresetNames() {
        var presetMap = loadPresetMap();
        var presetNames = [];
        for (var presetName in presetMap) {
            if (presetMap.hasOwnProperty(presetName)) presetNames.push(presetName);
        }
        presetNames.sort();
        return presetNames;
    }

    /**
     * 今のチェックと明るさを集める（状態を読めない項目・切り替えられない明るさは入れない）
     * @param {Object} dialogState - buildDialog() の戻り値
     * @returns {Object} 項目のキーをキーにした真偽と、uiBrightness に明るさの段階のキー
     */
    function collectPresetData(dialogState) {
        var presetData = {};
        var toggleRows = dialogState.toggleRows;
        for (var i = 0; i < toggleRows.length; i++) {
            if (toggleRows[i].state) presetData[toggleRows[i].item.key] = toggleRows[i].checkbox.value;
        }
        var brightnessPicker = dialogState.brightnessPicker;
        if (brightnessPicker && brightnessPicker.isEnabled) presetData.uiBrightness = BRIGHTNESS_LEVELS[brightnessPicker.selectedIndex].key;
        return presetData;
    }

    /**
     * プリセットのチェックと明るさを写す（プリセットに無い項目・状態を読めない項目はそのまま）
     * @param {Object} dialogState - buildDialog() の戻り値
     * @param {Object} presetData - collectPresetData() で集めたもの
     * @returns {void}
     */
    function applyPresetData(dialogState, presetData) {
        var brightnessPicker = dialogState.brightnessPicker;
        if (brightnessPicker && brightnessPicker.isEnabled) {
            for (var j = 0; j < BRIGHTNESS_LEVELS.length; j++) {
                if (BRIGHTNESS_LEVELS[j].key === presetData.uiBrightness) brightnessPicker.select(j);
            }
        }
        var toggleRows = dialogState.toggleRows;
        for (var i = 0; i < toggleRows.length; i++) {
            var toggleRow = toggleRows[i];
            if (toggleRow.state && typeof presetData[toggleRow.item.key] === "boolean") {
                toggleRow.checkbox.value = presetData[toggleRow.item.key];
            }
        }
    }

    /**
     * 最上部のプリセットの行（ドロップダウン・保存・削除）を作る
     * @param {Window} dialog - ダイアログボックス
     * @returns {{presetDropdown: DropDownList, presetSaveIcon: Button, presetDeleteIcon: Button}} プリセットの行のコントロール
     */
    function buildPresetRow(dialog) {
        var presetRow = dialog.add("group");
        setupRow(presetRow, "center", ROW_SPACING); /* ダイアログボックスの左右中央に置く / centered in the dialog */
        presetRow.add("statictext", undefined, labelText("fieldLabel.preset"));
        var presetDropdown = presetRow.add("dropdownlist", undefined, []);
        presetDropdown.helpTip = getLabel("tooltip.preset");
        presetDropdown.preferredSize.width = PRESET_DROPDOWN_WIDTH;
        /* 保存アイコンは塗りの面が大きく濃く見えるので、色を薄めて描く / the save icon is mostly solid and looks heavy, so draw it lighter */
        var presetSaveIcon = addIconButton(presetRow, PRESET_ICON_SIZE, function (iconGraphics, iconWidth, iconHeight, iconColor) {
            drawSaveIcon(iconGraphics, iconWidth, iconHeight, [iconColor[0], iconColor[1], iconColor[2], iconColor[3] * SAVE_ICON_OPACITY]);
        });
        presetSaveIcon.helpTip = getLabel("tooltip.presetSave");
        var presetDeleteIcon = addIconButton(presetRow, PRESET_ICON_SIZE, drawTrashIcon);
        presetDeleteIcon.helpTip = getLabel("tooltip.presetDelete");
        return { presetDropdown: presetDropdown, presetSaveIcon: presetSaveIcon, presetDeleteIcon: presetDeleteIcon };
    }

    /**
     * プリセットのドロップダウンを作り直し、指定の名前を選ぶ
     * @param {Object} presetControls - buildPresetRow() の戻り値
     * @param {string|null} selectedName - 選んでおくプリセット名（null なら「---」）
     * @returns {void}
     */
    function fillPresetDropdown(presetControls, selectedName) {
        var presetDropdown = presetControls.presetDropdown;
        var presetNames = getPresetNames();
        var selectedIndex = 0;
        presetDropdown.removeAll();
        presetDropdown.add("item", getLabel("dropdown.presetPlaceholder"));
        for (var i = 0; i < presetNames.length; i++) {
            presetDropdown.add("item", presetNames[i]);
            if (presetNames[i] === selectedName) selectedIndex = i + 1;
        }
        presetDropdown.selection = selectedIndex;
        setIconButtonEnabled(presetControls.presetDeleteIcon, selectedIndex > 0);
    }

    /**
     * 選んでいるプリセット名を返す
     * @param {Object} presetControls - buildPresetRow() の戻り値
     * @returns {string|null} プリセット名。「---」なら null
     */
    function getSelectedPresetName(presetControls) {
        var presetSelection = presetControls.presetDropdown.selection;
        return (presetSelection && presetSelection.index > 0) ? presetSelection.text : null;
    }

    /**
     * プリセット名を尋ねる
     * @param {string} initialName - 入力欄に入れておく名前
     * @returns {string|null} プリセット名。キャンセルか空なら null
     */
    function showPresetNameDialog(initialName) {
        var nameDialog = new Window("dialog", getLabel("dialog.presetSave"));
        setupWindow(nameDialog);
        var nameRow = nameDialog.add("group");
        setupRow(nameRow, "left", ROW_SPACING);
        nameRow.add("statictext", undefined, labelText("fieldLabel.presetName"));
        var nameInput = nameRow.add("edittext", undefined, initialName);
        nameInput.helpTip = getLabel("tooltip.presetName");
        nameInput.characters = PRESET_NAME_CHARS;
        nameInput.active = true;

        var buttonRow = addButtonRow(nameDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(nameDialog, SCRIPT_NAME + "_presetName");
        if (nameDialog.show() !== 1) return null;

        var presetName = nameInput.text.replace(/^\s+|\s+$/g, "");
        if (!presetName || presetName === getLabel("dropdown.presetPlaceholder")) return null;
        return presetName;
    }

    /**
     * プリセットの集まりを保存する。失敗したら知らせる
     * @param {Object} presetMap - プリセット名をキーにしたチェックの組み合わせ
     * @returns {boolean} 保存できたら true
     */
    function savePresetMap(presetMap) {
        if (presetSettingsStore.save(presetMap)) return true;
        alert(getLabel("alert.presetSaveFailed"));
        return false;
    }

    /**
     * 今のチェックを名前を付けて保存する（同じ名前は確認してから上書き）
     * @param {Object} presetControls - buildPresetRow() の戻り値
     * @param {Object} dialogState - buildDialog() の戻り値
     * @returns {string|null} 保存したプリセット名。保存しなかったら null
     */
    function saveCurrentPreset(presetControls, dialogState) {
        var presetName = showPresetNameDialog(getSelectedPresetName(presetControls) || "");
        if (!presetName) return null;
        var presetMap = loadPresetMap();
        if (presetMap.hasOwnProperty(presetName) && !confirm(getLabel("confirm.presetOverwrite", { name: presetName }), true)) return null;
        presetMap[presetName] = collectPresetData(dialogState);
        return savePresetMap(presetMap) ? presetName : null;
    }

    /**
     * 選んでいるプリセットを確認してから削除する
     * @param {Object} presetControls - buildPresetRow() の戻り値
     * @returns {boolean} 削除したら true
     */
    function deleteSelectedPreset(presetControls) {
        var presetName = getSelectedPresetName(presetControls);
        if (!presetName || !confirm(getLabel("confirm.presetDelete", { name: presetName }), true)) return false;
        var presetMap = loadPresetMap();
        delete presetMap[presetName];
        return savePresetMap(presetMap);
    }

    /**
     * プリセットの行のイベントをつなぐ。選んだら読み込む。
     * ドロップダウンの作り直しでも onChange が来るので、作り直している間は読まない
     * @param {Object} presetControls - buildPresetRow() の戻り値
     * @param {Object} dialogState - buildDialog() の戻り値
     * @returns {void}
     */
    function bindPresetEvents(presetControls, dialogState) {
        var isRefilling = false;

        /* ドロップダウンを作り直して名前を選ぶ（onChange は無視させる）/ refill the dropdown, ignoring its onChange */
        function refreshPresetDropdown(selectedName) {
            isRefilling = true;
            fillPresetDropdown(presetControls, selectedName);
            isRefilling = false;
        }

        presetControls.presetDropdown.onChange = function () {
            if (isRefilling) return;
            var presetName = getSelectedPresetName(presetControls);
            setIconButtonEnabled(presetControls.presetDeleteIcon, presetName !== null);
            var presetData = presetName ? loadPresetMap()[presetName] : null;
            if (presetData) applyPresetData(dialogState, presetData);
        };
        presetControls.presetSaveIcon.onClick = function () {
            var savedName = saveCurrentPreset(presetControls, dialogState);
            if (savedName) refreshPresetDropdown(savedName);
        };
        presetControls.presetDeleteIcon.onClick = function () {
            if (deleteSelectedPreset(presetControls)) refreshPresetDropdown(null);
        };
        refreshPresetDropdown(null);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 項目のメニューの道筋を、今の UI 言語で返す
     * @param {Object} toggleItem - VIEW_TOGGLE_GROUPS の項目
     * @returns {string} メニューの道筋
     */
    function getMenuPath(toggleItem) {
        return toggleItem.menu[uiLang] || toggleItem.menu.en;
    }

    /**
     * ツールバーの種類ごとのメニューの道筋を返す
     * @param {Object} toolbarKind - TOOLBAR_KINDS の要素
     * @returns {string} メニューの道筋
     */
    function getToolbarMenuPath(toolbarKind) {
        return (uiLang === "ja" ? "ウィンドウ>ツールバー>" : "Window>Toolbars>") + toolbarKind.menuName[uiLang];
    }

    /**
     * 種類名からツールバーの定義を返す
     * @param {string} kindName - "quick" / "basic" / "advanced"
     * @returns {Object|null} TOOLBAR_KINDS の要素。見つからなければ null
     */
    function findToolbarKind(kindName) {
        for (var i = 0; i < TOOLBAR_KINDS.length; i++) {
            if (TOOLBAR_KINDS[i].kind === kindName) return TOOLBAR_KINDS[i];
        }
        return null;
    }

    /**
     * 補助アプリで読む必要があるメニューの道筋を集める
     * @returns {string[]} メニューの道筋
     */
    function collectMenuPathsToRead() {
        var menuPaths = [];
        for (var i = 0; i < VIEW_TOGGLE_GROUPS.length; i++) {
            var groupItems = VIEW_TOGGLE_GROUPS[i].items;
            for (var j = 0; j < groupItems.length; j++) {
                var toggleItem = groupItems[j];
                if (toggleItem.toolbar) {
                    for (var k = 0; k < TOOLBAR_KINDS.length; k++) menuPaths.push(getToolbarMenuPath(TOOLBAR_KINDS[k]));
                } else if (toggleItem.menu && !toggleItem.pref) {
                    menuPaths.push(getMenuPath(toggleItem));
                }
            }
        }
        return menuPaths;
    }

    /**
     * 項目の今の状態を返す
     * @param {Object} toggleItem - VIEW_TOGGLE_GROUPS の項目
     * @param {Object|null} menuStates - getAiMenuState() の結果
     * @returns {{isOn: boolean, toolbarKind: Object|null}|null} 状態。読めなければ null
     */
    function readToggleState(toggleItem, menuStates) {
        /* 補助アプリで切り替える項目に、今の UI 言語の道筋が無ければ使えない / Unusable when a helper-switched item has no path for this UI language */
        if (toggleItem.menu && !toggleItem.command && !toggleItem.writePref && !getMenuPath(toggleItem)) return null;
        if (toggleItem.pref) {
            var isOn = (toggleItem.prefType === "integer") ? app.preferences.getIntegerPreference(toggleItem.pref) !== 0 : app.preferences.getBooleanPreference(toggleItem.pref);
            return { isOn: isOn, toolbarKind: null };
        }
        if (!menuStates) return null;
        if (toggleItem.toolbar) {
            /* どれかのツールバーに ✓ が付いていれば表示中 / Shown when any toolbar is checked */
            var foundAny = false;
            for (var i = 0; i < TOOLBAR_KINDS.length; i++) {
                var toolbarState = menuStates[getToolbarMenuPath(TOOLBAR_KINDS[i])];
                if (toolbarState === "on") return { isOn: true, toolbarKind: TOOLBAR_KINDS[i] };
                if (toolbarState === "off") foundAny = true;
            }
            return foundAny ? { isOn: false, toolbarKind: null } : null;
        }
        var menuState = menuStates[getMenuPath(toggleItem)];
        if (menuState !== "on" && menuState !== "off") return null;
        return { isOn: menuState === "on", toolbarKind: null };
    }

    /**
     * 項目を指定の状態に切り替える。コマンド ID があればすぐ実行し、無ければ補助アプリへの依頼に積む
     * @param {Object} toggleItem - VIEW_TOGGLE_GROUPS の項目
     * @param {{isOn: boolean, toolbarKind: Object|null}} currentState - 今の状態
     * @param {boolean} turnOn - 入れるなら true
     * @param {Array} helperRequests - 補助アプリへの依頼（[道筋, "on"|"off"]）を積む配列
     * @returns {void}
     */
    function applyToggle(toggleItem, currentState, turnOn, helperRequests) {
        if (toggleItem.toolbar) {
            /* 隠すときは表示中のツールバー、出すときは TOOLBAR_TO_SHOW のツールバーを、補助アプリが切り替える
               The helper hides the toolbar that is shown, or shows the one in TOOLBAR_TO_SHOW */
            var targetKind = turnOn ? findToolbarKind(TOOLBAR_TO_SHOW) : currentState.toolbarKind;
            if (targetKind) helperRequests.push([getToolbarMenuPath(targetKind), turnOn ? "on" : "off"]);
            return;
        }
        if (toggleItem.command) {
            app.executeMenuCommand(toggleItem.command);
            return;
        }
        if (toggleItem.writePref) {
            app.preferences.setBooleanPreference(toggleItem.pref, turnOn);
            return;
        }
        helperRequests.push([getMenuPath(toggleItem), turnOn ? "on" : "off"]);
    }

    /**
     * 縦に積む列を横に並べて作る
     * @param {Window} dialog - ダイアログボックス
     * @param {number} columnCount - 列の数
     * @returns {Group[]} 左から順の列
     */
    function addColumnGroups(dialog, columnCount) {
        var columnsRow = dialog.add("group");
        columnsRow.orientation = "row";
        columnsRow.alignChildren = ["fill", "top"];
        columnsRow.spacing = COLUMN_SPACING;
        var columnGroups = [];
        for (var i = 0; i < columnCount; i++) {
            var columnGroup = columnsRow.add("group");
            columnGroup.orientation = "column";
            columnGroup.alignChildren = ["fill", "top"];
            columnGroup.spacing = WINDOW_SPACING;
            columnGroups.push(columnGroup);
        }
        return columnGroups;
    }

    /**
     * カテゴリのパネルを作り、項目ごとのチェックボックスを今の状態で並べる
     * @param {Group} parentColumn - パネルを足す列
     * @param {Object} toggleGroup - VIEW_TOGGLE_GROUPS の要素
     * @param {Object|null} menuStates - getAiMenuState() の結果
     * @param {Object} dialogState - 項目・状態・チェックボックスの組（toggleRows）と明るさの見本（brightnessPicker）を入れる
     * @returns {void}
     */
    function addTogglePanel(parentColumn, toggleGroup, menuStates, dialogState) {
        var togglePanel = parentColumn.add("panel", undefined, getLabel("panel." + toggleGroup.labelKey));
        setupPanel(togglePanel, 6);
        /* 明るさは環境設定のボタンで切り替えるので、補助アプリが使えて日本語版のときだけ選べる
           Brightness is switched with the buttons in Preferences, so it needs the helper and the Japanese UI */
        if (toggleGroup.hasBrightness) dialogState.brightnessPicker = addBrightnessPicker(togglePanel, menuStates !== null && uiLang === "ja");
        for (var i = 0; i < toggleGroup.items.length; i++) {
            var toggleItem = toggleGroup.items[i];
            var currentState = readToggleState(toggleItem, menuStates);
            var toggleCheckbox = togglePanel.add("checkbox", undefined, getLabel("checkbox." + toggleItem.key));
            if (currentState) {
                toggleCheckbox.value = currentState.isOn;
                if (LABELS.tooltip[toggleItem.key]) toggleCheckbox.helpTip = getLabel("tooltip." + toggleItem.key, { toolbarName: findToolbarKind(TOOLBAR_TO_SHOW).menuName[uiLang] });
            } else {
                toggleCheckbox.enabled = false;
                toggleCheckbox.helpTip = getLabel("tooltip.unavailable");
            }
            dialogState.toggleRows.push({ item: toggleItem, state: currentState, checkbox: toggleCheckbox });
        }
    }

    /**
     * ダイアログボックスを作る。左の列に基本UI・オブジェクト、右の列にガイド・グリッド・その他のスナップを置く
     * @param {Object|null} menuStates - getAiMenuState() の結果
     * @returns {{dialog: Window, toggleRows: Array, brightnessPicker: Object|null}} ダイアログボックス、項目・状態・チェックボックスの組、明るさの見本
     */
    function buildDialog(menuStates) {
        var dialog = new Window("dialog");
        dialog.text = getLabel("dialog.title") + " " + SCRIPT_VERSION;
        setupWindow(dialog);

        var presetControls = buildPresetRow(dialog);
        var columnGroups = addColumnGroups(dialog, 2);
        var dialogState = { dialog: dialog, toggleRows: [], brightnessPicker: null };
        for (var i = 0; i < VIEW_TOGGLE_GROUPS.length; i++) {
            addTogglePanel(columnGroups[i < LEFT_COLUMN_GROUP_COUNT ? 0 : 1], VIEW_TOGGLE_GROUPS[i], menuStates, dialogState);
        }

        var buttonRow = addButtonRow(dialog);
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        bindPresetEvents(presetControls, dialogState);
        return dialogState;
    }

    /**
     * メイン処理：状態を読み、ダイアログボックスで選んだとおりに切り替える
     * @returns {void}
     */
    function main() {
        var menuStates = getAiMenuState(collectMenuPathsToRead());
        if (!menuStates) alert(getLabel("alert.helperUnavailable"));

        var dialogState = buildDialog(menuStates);
        prepareDialogWindow(dialogState.dialog, SCRIPT_NAME);
        if (dialogState.dialog.show() !== 1) return;

        var helperRequests = [];
        var hasWrittenPref = false;
        var toggleRows = dialogState.toggleRows;
        for (var i = 0; i < toggleRows.length; i++) {
            var toggleRow = toggleRows[i];
            if (!toggleRow.state || toggleRow.checkbox.value === toggleRow.state.isOn) continue;
            applyToggle(toggleRow.item, toggleRow.state, toggleRow.checkbox.value, helperRequests);
            if (toggleRow.item.writePref) hasWrittenPref = true;
        }
        /* 設定キーを書いただけでは描き直されないので、ズームを往復させる / Writing the key does not redraw, so zoom out and back in */
        if (hasWrittenPref && app.documents.length) {
            app.executeMenuCommand("zoomout");
            app.executeMenuCommand("zoomin");
        }
        var brightnessPicker = dialogState.brightnessPicker;
        if (brightnessPicker && brightnessPicker.selectedIndex !== brightnessPicker.initialIndex) {
            helperRequests.push(["環境設定>ユーザーインターフェイス...>" + BRIGHTNESS_LEVELS[brightnessPicker.selectedIndex].prefButton, "on"]);
        }
        /* コマンド ID の無い項目は、スクリプトの終了後に補助アプリが切り替える / Items without a command ID are switched by the helper after the script ends */
        if (helperRequests.length) setAiMenuState(helperRequests);
    }

    main();

})();
