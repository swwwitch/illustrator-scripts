#target illustrator
#targetengine "EditGradientStopsEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.EditGradientStops
Name=グラデーションストップを一覧表示して編集
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.1
Release-Date=2026-11-20
Description-BEGIN
選択したオブジェクトに使われているグラデーションのストップを一覧表示し、カラー・位置・不透明度の変更や、ストップどうしの入れ替えができます。
指定した％の位置で色が切り替わるセパレートグラデーションへの変換や、グラデーションの無いオブジェクトの色からの作成もできます。

Lists the gradient stops used by the selected objects and lets you change their color, location and opacity, or swap stops.
It can also turn a gradient into a segmented one whose colors switch at the percentages you enter, or build one from the colors of objects without gradients.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択したオブジェクトに使われているグラデーションのストップを一覧表示し、カラー・位置・不透明度の変更や、ストップどうしの入れ替えができます。
指定した％の位置で色が切り替わるセパレートグラデーションへの変換や、グラデーションの無いオブジェクトの色からの作成もできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/EditGradientStops.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n524fafa198dd

### Overview

Lists the gradient stops used by the selected objects and lets you change their color, location and opacity, or swap stops.
It can also turn a gradient into a segmented one whose colors switch at the percentages you enter, or build one from the colors of objects without gradients.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/EditGradientStops.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "EditGradientStops";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-20";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "グラデーションストップを一覧表示して編集"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/EditGradientStops.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/EditGradientStops.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n524fafa198dd"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/EditGradientStops.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 一覧に並べるグラデーションの上限（超えた分は件数だけ表示） / Max gradients listed (the rest are only counted) */
    var MAX_LISTED_GRADIENTS = 10;

    /* グラデーションに残すストップの最小数（Illustrator は2つ未満にできない） / Minimum stops a gradient keeps (Illustrator needs two) */
    var MIN_STOP_COUNT = 2;

    /* 位置・不透明度の欄の増減量・範囲・単位 / Step, range and unit of the location and opacity fields */
    var PERCENT_FIELD_OPTIONS = { step: 1, min: 0, max: 100, unit: "%" };

    /* セパレートで扱える最大の色数（境界の欄は色数より1つ少ない） / Max colors for a segmented gradient (one boundary field fewer) */
    var MAX_SEPARATE_COLORS = 10;

    /* 境界の前後に置く2つのストップのずれ（%）。同じ位置には重ねず、境界の±この値に置く / Offset of the stop pair on each side of a boundary (%) */
    var BOUNDARY_OFFSET = 0.01;

    /* 色の切り替わりの幅がこれ以内（%）ならセパレートとみなし、今の境界を初期値にする / Max color transition width (%) treated as a hard edge */
    var SEPARATE_EDGE_MAX_GAP = 1;

    /* 境界の欄の増減量・範囲・単位 / Step, range and unit of the boundary fields */
    var BOUNDARY_FIELD_OPTIONS = { step: 1, min: 0, max: 100, unit: "%" };

    /* 新規作成するグラデーションの名前 / Name of a new gradient */
    var NEW_GRADIENT_BASE_NAME = "Separate Gradient";

    /* 新規作成で、選択の大きさが測れないときの長方形の大きさ（pt） / Rectangle size when the selection cannot be measured (pt) */
    var DEFAULT_RECT_SIZE = 100;

    /* ［丸型の破線］で入れる破線（pt）。[線分] だけなら間隔は空白（線分と同じ間隔） / Dash pattern set by Round Dashes (pt); a single value leaves the gap blank */
    var STROKE_DASH_PATTERN = [1];

    // =========================================
    // レイアウト / Layout
    // =========================================

    var COLOR_CHIP_SIZE    = 14;  /* カラーの見本の大きさ / color chip size */
    var COLOR_TEXT_WIDTH   = 180; /* カラーの値の欄の幅 / width of the color value column */
    var VALUE_COLUMN_WIDTH = 90;  /* 位置・不透明度の欄の幅（∧∨と入力欄が収まる幅） / width of the location and opacity columns, stepper included */
    var PERCENT_FIELD_CHARACTERS = 5; /* 位置・不透明度の入力欄の文字数 / characters of the location and opacity fields */
    var STOP_ROW_SPACING   = 6;   /* ストップの行の間隔 / spacing between stop rows */
    var STOP_CELL_SPACING  = 8;   /* 行の中の欄の間隔 / spacing between cells in a row */
    var HEADER_LABEL_INSET = 3;   /* 位置・不透明度の見出しを入力欄の左端から右へずらす量（欄の中の文字に揃える） / header offset from the field's left edge, to line up with the text inside */
    var ROW_BUTTON_SIZE    = [18, 22]; /* 行の末尾のボタン（入れ替え・複製）の大きさ / size of the row buttons (swap, duplicate) */
    var ROW_BUTTON_SPACING = 0;   /* 行の末尾のボタンの間隔 / spacing between the row buttons */
    var ROW_BUTTON_GROUP_LEFT_MARGIN = 6; /* 不透明度の欄とボタンの間に足す余白 / extra gap between the opacity field and the row buttons */
    var DUPLICATE_ICON_SQUARE_SIZE = 8;   /* 複製のアイコンの四角の大きさ / square size of the duplicate icon */
    var DUPLICATE_ICON_SHIFT       = 3;   /* 複製のアイコンの2つの四角のずれ / offset between the two squares */
    var DUPLICATE_ICON_LINE_WIDTH  = 1.2; /* 複製のアイコンの線の太さ / line width of the duplicate icon */
    var REMOVE_ICON_SIZE           = 7;   /* 削除のアイコン（×）の大きさ / size of the remove (x) icon */
    var REMOVE_ICON_LINE_WIDTH     = 1.4; /* 削除のアイコンの線の太さ / line width of the remove icon */
    var SWAP_ARROW_LENGTH  = 11;  /* 矢印の長さ / arrow length */
    var SWAP_ARROW_HEAD_SIZE = 3.5; /* 矢じりの大きさ / arrowhead size */
    var SWAP_ARROW_LINE_WIDTH = 1.4; /* 矢印の線の太さ / arrow line width */
    var BOUNDARY_LABEL_WIDTH = 60;    /* 境界の欄の項目名の幅 / label width of the boundary fields */
    var BOUNDARY_FIELD_CHARACTERS = 6; /* 境界の入力欄の文字数 / characters of the boundary fields */

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
    var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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

    /**
     * 自動計算アイコン（魔法の杖と3つのきらめき）を描く。24×22 の絵を、縦横比を保って中央に置く。
     * 行数・列数や比率など、値を自動で求める［自動］ボタンの代わりに使う
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawAutoIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        var iconScale = Math.min(iconWidth / 24, iconHeight / 22);
        var originX = (iconWidth - 24 * iconScale) / 2;
        var originY = (iconHeight - 22 * iconScale) / 2;

        /** 絵の座標で線を1本引く / stroke one line in design coordinates */
        function strokeLine(x1, y1, x2, y2, lineWidth) {
            iconGraphics.newPath();
            iconGraphics.moveTo(originX + x1 * iconScale, originY + y1 * iconScale);
            iconGraphics.lineTo(originX + x2 * iconScale, originY + y2 * iconScale);
            iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, lineWidth * iconScale));
        }

        /** 十字のきらめきを描く / draw a cross-shaped sparkle */
        function drawSparkle(centerX, centerY, radius) {
            strokeLine(centerX - radius, centerY, centerX + radius, centerY, 1.2);
            strokeLine(centerX, centerY - radius, centerX, centerY + radius, 1.2);
        }

        strokeLine(4, 19, 14, 9, 2.2);     /* 杖 / wand */
        strokeLine(14, 9, 16.5, 6.5, 3.2); /* 杖の先 / wand tip */
        drawSparkle(18.5, 12.5, 2.5);
        drawSparkle(9, 5, 2.5);
        drawSparkle(20, 3.5, 1.8);
    }

    // アイコンのボタン（再利用パーツ）ここまで / End of the reusable icon buttons

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

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
       9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
       / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
         nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
    var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
    var VERSION_FOOTER_GLYPHS = {
        "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
        "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
        "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
        "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
        "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
        "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
        "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
        "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
        "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
        "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
        "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
        "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
        "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
        "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
        "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
        "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
        "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
        "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
        "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
        "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
        "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
        "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
        "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
        "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
        "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
        "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
        "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
        "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
        "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
        "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
        "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
        "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
        "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
        "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
        "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
        "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
        "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
        "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
        "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
        ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
        "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
        "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
        "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
        " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
    };

    /**
     * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPT_VERSION）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

        /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
           左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
           / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
             or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
        var versionText = String(SCRIPT_VERSION);
        var leftCandidates = [versionText];
        if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
        var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
        versionCanvas.alignment = ["fill", "center"];
        bindVersionInfoClick(versionCanvas); /* クリックで「このスクリプトについて」を開く / click opens About This Script */
        var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
        brandCanvas.alignment = ["right", "center"];
        brandCanvas.maximumSize = brandCanvas.preferredSize;
        targetWindow.versionFooter = footerGroup;
        return versionCanvas;
    }

    /**
     * 文字列のすべての文字に字形があるかを返す
     * @param {string} dotText - 文字列
     * @returns {boolean} すべてあれば true
     */
    function hasVersionFooterGlyphs(dotText) {
        for (var i = 0; i < dotText.length; i++) {
            if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
        }
        return dotText.length > 0;
    }

    /**
     * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
     * @param {Group} parentGroup - 足す先
     * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
     * @returns {Group} 描くグループ
     */
    function addVersionFooterCanvas(parentGroup, dotTexts) {
        var dotCanvas = parentGroup.add("group");
        var shortestText = dotTexts[dotTexts.length - 1];
        dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
        dotCanvas.helpTip = dotTexts[0];
        dotCanvas.onDraw = function () {
            var availableWidth = this.size ? this.size.width : 0;
            var drawText = shortestText;
            for (var i = 0; i < dotTexts.length; i++) {
                if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                    drawText = dotTexts[i];
                    break;
                }
            }
            drawVersionFooterDots(this.graphics, drawText);
        };
        return dotCanvas;
    }

    /**
     * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - 文字列
     * @returns {number} 幅（ドット）
     */
    function measureVersionFooterDots(versionText) {
        var dotCount = 0;
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
            dotCount += glyphRows[0].length;
        }
        return dotCount;
    }

    /**
     * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - 文字列
     * @returns {void}
     */
    function drawVersionFooterDots(canvasGraphics, versionText) {
        var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
        var dotSize = VERSION_FOOTER_PIXEL_SIZE;
        var cursorDots = 0;
        canvasGraphics.newPath();
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
            for (var row = 0; row < glyphRows.length; row++) {
                var rowText = glyphRows[row];
                var runStart = -1;
                for (var col = 0; col <= rowText.length; col++) {
                    var isDot = col < rowText.length && rowText.charAt(col) === "#";
                    if (isDot && runStart < 0) runStart = col;
                    if (!isDot && runStart >= 0) {
                        canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                        runStart = -1;
                    }
                }
            }
            cursorDots += glyphRows[0].length;
        }
        canvasGraphics.fillPath(footerBrush);
    }

    // -----------------------------------------
    // このスクリプトについて / About this script
    // -----------------------------------------
    /* 左下のスクリプト名・バージョンをクリックすると開く。基本情報ブロックの SCRIPT_* を読み、
       この部品だけで動くよう、文言（VERSION_INFO_LABELS）とレイアウトの値も部品の中に持つ
       Opened by clicking the name and version at the bottom left. Reads the SCRIPT_* basic info,
       and keeps its own labels and layout values so the part works on its own */

    var VERSION_INFO_LANG = (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en"; /* 文言の言語 / label language */
    var VERSION_INFO_IS_INDESIGN = (String(app.name).indexOf("InDesign") >= 0);         /* InDesign で動いているか / running in InDesign */
    /* 実行中のスクリプトのファイル。ボタンを押すころには $.fileName が別のものを指すことがあるので、読み込んだ時点で控える
       The running script file, captured at load time; $.fileName may point elsewhere by the time the button is clicked */
    var VERSION_INFO_SCRIPT_FILE = new File($.fileName);
    var VERSION_INFO_REVEAL_APP_PATH = "/Applications/RevealInFinder.app";   /* Finder で選択表示する補助アプリ / helper app that reveals a file in the Finder */
    var VERSION_INFO_REVEAL_PATH_FILE = "/tmp/illustrator_reveal_path.txt";  /* 補助アプリが読むパス（アプリの旧名のまま。変えると読めない）/ path the helper reads (keeps the app's old name) */

    var VERSION_INFO_WINDOW_MARGINS = 16;               /* ウィンドウの余白 / window margin */
    var VERSION_INFO_WINDOW_SPACING = 12;               /* ウィンドウ内の間隔 / window spacing */
    var VERSION_INFO_PANEL_MARGINS = [16, 20, 16, 12];  /* パネルの余白 [左,上,右,下] / panel margins */
    var VERSION_INFO_ROW_SPACING = 6;                   /* 項目名と値の間隔 / gap between a label and its value */
    var VERSION_INFO_BUTTON_SPACING = 10;               /* ボタンどうしの間隔 / gap between buttons */
    var VERSION_INFO_BUTTON_ROW_MARGINS = [0, 5, 0, 14]; /* ボタン行の余白（Illustrator 標準のダイアログに合わせる）/ button row margins */
    var VERSION_INFO_COPY_BUTTON_SIZE = 22;             /* コピーボタンの一辺 / copy button size */
    /* コピーボタンの配色。∧∨（ステップボタン）と同じく、背景に重ねる半透明 / Copy button colors, translucent overlays like the stepper */
    var VERSION_INFO_COPY_FILL_COLOR    = isDarkUI() ? [0, 0, 0, 0.10] : [1, 1, 1, 0.50]; /* 地 / background */
    var VERSION_INFO_COPY_FRAME_COLOR   = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.15]; /* 枠線 / frame */
    var VERSION_INFO_COPY_PRESSED_COLOR = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 押下中 / pressed */
    var VERSION_INFO_COPY_ICON_COLOR    = isDarkUI() ? [1, 1, 1, 0.85] : [0, 0, 0, 0.70]; /* アイコン / icon */

    var VERSION_INFO_LABELS = {
        dialog: {
            title: { ja: "このスクリプトについて", en: "About This Script" }
        },
        panel: {
            generalInfo: { ja: "基本情報", en: "General" },
            links: { ja: "リンク", en: "Links" }
        },
        fieldLabel: {
            scriptName: { ja: "スクリプト名", en: "Script Name" },
            scriptFeature: { ja: "機能", en: "Feature" },
            scriptVersion: { ja: "バージョン", en: "Version" }
        },
        button: {
            openReadme: { ja: "README", en: "README" },
            openSourceCode: { ja: "ソースコード", en: "Source Code" },
            openArticle: { ja: "記事（note）", en: "Article (note)" },
            revealInFinder: { ja: "Finder に表示", en: "Show in Finder" },
            revealInExplorer: { ja: "エクスプローラーに表示", en: "Show in Explorer" },
            close: { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            versionInfo: { ja: "クリックで、このスクリプトについての情報を表示", en: "Click for info about this script" },
            copyScriptName: { ja: "スクリプト名をコピー", en: "Copy Script Name" }
        },
        alert: {
            copyFailed: { ja: "スクリプト名をコピーできませんでした。", en: "Could not copy the script name." }
        }
    };

    /**
     * VERSION_INFO_LABELS の文言を今の言語で返す
     * @param {Object} labelEntry - { ja, en }
     * @returns {string} 文言
     */
    function getVersionInfoLabel(labelEntry) {
        return labelEntry[VERSION_INFO_LANG] || labelEntry.en;
    }

    /**
     * 基本情報ブロックの値を返す（変数が無い・文字列でなければ空文字）。typeof で読むので、未定義でも例外にならない
     * @param {string} variableKey - "name" / "version" / "feature" / "readme" / "source" / "article"
     * @returns {string} 値
     */
    function getVersionInfoValue(variableKey) {
        var infoValue;
        if (variableKey === "name") infoValue = (typeof SCRIPT_NAME !== "undefined") ? SCRIPT_NAME : "";
        else if (variableKey === "version") infoValue = (typeof SCRIPT_VERSION !== "undefined") ? SCRIPT_VERSION : "";
        else if (variableKey === "feature") infoValue = (typeof SCRIPT_DESCRIPTION_JA !== "undefined") ? SCRIPT_DESCRIPTION_JA : "";
        else if (variableKey === "readme") {
            if (VERSION_INFO_LANG === "ja") infoValue = (typeof SCRIPT_README_JA !== "undefined") ? SCRIPT_README_JA : "";
            else infoValue = (typeof SCRIPT_README_EN !== "undefined") ? SCRIPT_README_EN : "";
        }
        else if (variableKey === "source") infoValue = (typeof SCRIPT_SOURCE_URL !== "undefined") ? SCRIPT_SOURCE_URL : "";
        else if (variableKey === "article") infoValue = (typeof SCRIPT_ARTICLE_URL !== "undefined") ? SCRIPT_ARTICLE_URL : "";
        return (infoValue === undefined || infoValue === null) ? "" : String(infoValue);
    }

    /**
     * 左下のスクリプト名・バージョンをクリックしたら、「このスクリプトについて」のダイアログを開くようにする
     * @param {Group} versionCanvas - addVersionFooterCanvas() で作った左の欄
     * @returns {void}
     */
    function bindVersionInfoClick(versionCanvas) {
        versionCanvas.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.versionInfo);
        /* mousedown で開くと、離したときのイベントが閉じたあとの欄に届くので mouseup で開く
           Open on mouseup so the release does not land on the footer after the dialog closes */
        versionCanvas.addEventListener("mouseup", function () {
            showVersionInfoDialog(versionCanvas.window.type === "palette");
        });
    }

    /**
     * 「このスクリプトについて」のダイアログを開く。スクリプト名（コピーボタン付き）・機能・バージョンの［基本情報］パネル、
     * README・ソースコード・記事を開く［リンク］パネル、［Finder に表示］［閉じる］を並べる
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true（Illustrator のコピーを BridgeTalk 経由にする）
     * @returns {void}
     */
    function showVersionInfoDialog(isFromPalette) {
        var infoDialog = new Window("dialog", getVersionInfoLabel(VERSION_INFO_LABELS.dialog.title));
        infoDialog.orientation = "column";
        infoDialog.alignChildren = "fill";
        infoDialog.margins = VERSION_INFO_WINDOW_MARGINS;
        infoDialog.spacing = VERSION_INFO_WINDOW_SPACING;

        var infoPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.generalInfo);
        /* コピーボタンはスクリプト名だけ / Only the script name gets a copy button */
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptName, getVersionInfoValue("name"), isFromPalette);
        /* 機能は README の見出しで日本語だけなので、日本語の画面だけに出す / The feature is the Japanese README title, shown only in the Japanese UI */
        if (VERSION_INFO_LANG === "ja" && getVersionInfoValue("feature") !== "") {
            addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptFeature, getVersionInfoValue("feature"), null);
        }
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptVersion, getVersionInfoValue("version"), null);
        alignVersionInfoLabels(infoPanel);

        /* URL が無いボタンはディム表示にして、並びはスクリプトごとに変えない / Dim buttons without a URL; keep the same buttons everywhere */
        var linksPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.links);
        var linkButtonRow = linksPanel.add("group");
        linkButtonRow.orientation = "row";
        linkButtonRow.alignment = ["left", "center"];
        linkButtonRow.alignChildren = ["left", "center"];
        linkButtonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openReadme, getVersionInfoValue("readme"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openSourceCode, getVersionInfoValue("source"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openArticle, getVersionInfoValue("article"));

        /* ボタン行：左に［Finder に表示］、右に［閉じる］ / Button row: Show in Finder at the left, Close at the right */
        var buttonRow = infoDialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = ["fill", "bottom"];
        buttonRow.alignChildren = ["left", "center"];
        buttonRow.margins = VERSION_INFO_BUTTON_ROW_MARGINS;
        buttonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        var revealLabel = (File.fs === "Windows") ? VERSION_INFO_LABELS.button.revealInExplorer : VERSION_INFO_LABELS.button.revealInFinder;
        var revealButton = buttonRow.add("button", undefined, getVersionInfoLabel(revealLabel));
        if (VERSION_INFO_SCRIPT_FILE.exists) {
            revealButton.helpTip = VERSION_INFO_SCRIPT_FILE.fsName;
            revealButton.onClick = function () {
                revealVersionInfoScriptFile(VERSION_INFO_SCRIPT_FILE);
                infoDialog.close();
            };
        } else {
            revealButton.enabled = false;
        }
        /* 余りの幅を吸って、［閉じる］を右端に寄せる / Absorbs the extra width so Close sits at the right edge */
        var buttonSpacer = buttonRow.add("group");
        buttonSpacer.alignment = ["fill", "fill"];
        buttonSpacer.minimumSize.width = 0;
        var closeButton = buttonRow.add("button", undefined, getVersionInfoLabel(VERSION_INFO_LABELS.button.close), { name: "ok" });
        closeButton.alignment = ["right", "center"];
        /* return でも esc でも閉じる / Close with both Return and Esc */
        infoDialog.defaultElement = closeButton;
        infoDialog.cancelElement = closeButton;
        infoDialog.onShow = function () {
            pinVersionInfoCopyButton(infoPanel);
        };
        infoDialog.show();
    }

    /**
     * 「このスクリプトについて」のダイアログにパネルを足す
     * @param {Window} infoDialog - 足す先のダイアログ
     * @param {Object} titleEntry - パネル名の文言 { ja, en }
     * @returns {Panel} パネル
     */
    function addVersionInfoPanel(infoDialog, titleEntry) {
        var infoPanel = infoDialog.add("panel", undefined, getVersionInfoLabel(titleEntry));
        infoPanel.orientation = "column";
        infoPanel.alignChildren = ["fill", "top"];
        infoPanel.alignment = "fill";
        infoPanel.margins = VERSION_INFO_PANEL_MARGINS;
        infoPanel.spacing = VERSION_INFO_WINDOW_SPACING;
        return infoPanel;
    }

    /**
     * ［基本情報］パネルに「項目名 : 値」の1行を足す（値は見せるだけなので statictext）
     * @param {Panel} infoPanel - 足す先のパネル
     * @param {Object} labelEntry - 項目名の文言 { ja, en }
     * @param {string} valueText - 表示する値
     * @param {boolean|null} copyFromPalette - 値の右にコピーボタンを置くなら true / false（常駐パレットから開いたか）、置かないなら null
     * @returns {StaticText} 値の statictext
     */
    function addVersionInfoRow(infoPanel, labelEntry, valueText, copyFromPalette) {
        var infoRow = infoPanel.add("group");
        infoRow.orientation = "row";
        infoRow.alignment = ["fill", "center"];
        infoRow.alignChildren = ["left", "center"];
        infoRow.spacing = VERSION_INFO_ROW_SPACING;
        /* 項目名のコロンは日本語が「 :」、英語が「:」 / Colon: " :" in Japanese, ":" in English */
        var rowLabel = infoRow.add("statictext", undefined, getVersionInfoLabel(labelEntry) + (VERSION_INFO_LANG === "ja" ? " :" : ":"));
        rowLabel.justify = "right";
        var valueLabel = infoRow.add("statictext", undefined, valueText);
        if (copyFromPalette === null) return valueLabel;
        /* statictext の fill では伸びないことがあるので、伸びるスペーサーでコピーボタンをパネルの右端に寄せる
           A stretching spacer (statictext fill is unreliable) pushes the copy button to the panel's right edge */
        var copySpacer = infoRow.add("group");
        copySpacer.alignment = ["fill", "fill"];
        copySpacer.minimumSize.width = 0;
        infoRow.copyButton = addVersionInfoCopyButton(infoRow, valueLabel, copyFromPalette);
        return valueLabel;
    }

    /**
     * コピーボタンを［基本情報］パネルの右端に置き直す（表示時、レイアウトが決まったあとに呼ぶ）。
     * 行の fill・スペーサーだけでは、項目の値が短い行がパネルの幅まで伸びないことがあるので、行の幅とボタンの位置を直接入れる
     * @param {Panel} infoPanel - ［基本情報］パネル
     * @returns {void}
     */
    function pinVersionInfoCopyButton(infoPanel) {
        var innerWidth = infoPanel.size.width - VERSION_INFO_PANEL_MARGINS[0] - VERSION_INFO_PANEL_MARGINS[2];
        for (var i = 0; i < infoPanel.children.length; i++) {
            var infoRow = infoPanel.children[i];
            if (!infoRow.copyButton || !infoRow.size || infoRow.size.width >= innerWidth) continue;
            infoRow.size = [innerWidth, infoRow.size.height];
            infoRow.copyButton.location = [innerWidth - infoRow.copyButton.size.width, infoRow.copyButton.location[1]];
        }
    }

    /**
     * ［基本情報］パネルの項目名を、いちばん長いものの幅にそろえる
     * @param {Panel} infoPanel - 対象のパネル
     * @returns {void}
     */
    function alignVersionInfoLabels(infoPanel) {
        var rowLabels = [];
        var labelWidth = 0;
        for (var i = 0; i < infoPanel.children.length; i++) {
            var rowLabel = infoPanel.children[i].children[0]; /* 行の先頭は項目名 / each row starts with its label */
            rowLabels.push(rowLabel);
            labelWidth = Math.max(labelWidth, rowLabel.graphics.measureString(rowLabel.text)[0]);
        }
        for (var j = 0; j < rowLabels.length; j++) {
            rowLabels[j].preferredSize.width = Math.ceil(labelWidth) + 2;
        }
    }

    /**
     * URL を開くボタンを足す。URL が無ければディム表示にする。開いたらダイアログは閉じる
     * @param {Window} infoDialog - ボタンを置くダイアログ
     * @param {Group} buttonGroup - 足す先のグループ
     * @param {Object} labelEntry - ボタンの文言 { ja, en }
     * @param {string} targetUrl - 開く URL（空ならディム表示）
     * @returns {void}
     */
    function addVersionInfoLinkButton(infoDialog, buttonGroup, labelEntry, targetUrl) {
        var linkButton = buttonGroup.add("button", undefined, getVersionInfoLabel(labelEntry));
        if (!targetUrl) {
            linkButton.enabled = false;
            return;
        }
        linkButton.helpTip = targetUrl;
        linkButton.onClick = function () {
            openVersionInfoUrl(targetUrl);
            infoDialog.close();
        };
    }

    /**
     * 値をクリップボードへコピーするアイコンのボタンを足す。コピーしたらアイコンをチェックマークに変え、ポインタが外れたら戻す
     * @param {Group} infoRow - 足す先の行
     * @param {StaticText} valueLabel - コピーする値の statictext
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {Group} ボタンとして使う group
     */
    function addVersionInfoCopyButton(infoRow, valueLabel, isFromPalette) {
        var buttonSize = VERSION_INFO_COPY_BUTTON_SIZE;
        var copyButton = infoRow.add("group");
        copyButton.preferredSize = [buttonSize, buttonSize];
        copyButton.minimumSize = [buttonSize, buttonSize];
        copyButton.maximumSize = [buttonSize, buttonSize];
        copyButton.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.copyScriptName);
        copyButton.isPressed = false;
        copyButton.isCopied = false;

        copyButton.onDraw = function () {
            var buttonGraphics = copyButton.graphics;
            var fillColor = copyButton.isPressed ? VERSION_INFO_COPY_PRESSED_COLOR : VERSION_INFO_COPY_FILL_COLOR;
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0, 0, buttonSize, buttonSize);
            buttonGraphics.fillPath(buttonGraphics.newBrush(buttonGraphics.BrushType.SOLID_COLOR, fillColor));
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0.5, 0.5, buttonSize - 1, buttonSize - 1);
            buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_FRAME_COLOR, 1));
            if (copyButton.isCopied) drawVersionInfoCheckIcon(buttonGraphics, buttonSize);
            else drawVersionInfoCopyIcon(buttonGraphics, buttonSize);
        };

        /**
         * 状態を変えて描き直す（group には notify() が無いので隠して再表示する）
         * @param {boolean} isPressed - 押下中なら true
         * @param {boolean} isCopied - コピー済みなら true
         * @returns {void}
         */
        function repaint(isPressed, isCopied) {
            if (copyButton.isPressed === isPressed && copyButton.isCopied === isCopied) return;
            copyButton.isPressed = isPressed;
            copyButton.isCopied = isCopied;
            copyButton.hide();
            copyButton.show();
        }
        copyButton.addEventListener("mousedown", function () { repaint(true, false); });
        copyButton.addEventListener("mouseup", function () {
            if (!copyButton.isPressed) return;
            var hasCopied = copyVersionInfoText(valueLabel.text, isFromPalette);
            repaint(false, hasCopied);
            if (!hasCopied) alert(getVersionInfoLabel(VERSION_INFO_LABELS.alert.copyFailed));
        });
        /* 押したまま外へ出たとき・コピーのあと離れたときは元の見た目に戻す / Reset when the pointer leaves */
        copyButton.addEventListener("mouseout", function () { repaint(false, false); });
        return copyButton;
    }

    /**
     * コピーのアイコン（後ろの四角に前の四角を重ねた形）を描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCopyIcon(buttonGraphics, buttonSize) {
        var iconPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1);
        var sheetSize = 9;
        var offset = 3;
        var frontLeft = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var frontTop = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var backLeft = frontLeft - offset;
        var backTop = frontTop - offset;
        /* 後ろの四角は、前の四角に隠れる部分を描かない（上と左の辺だけ）/ Draw only the visible top and left of the back sheet */
        buttonGraphics.newPath();
        buttonGraphics.moveTo(frontLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, frontTop);
        buttonGraphics.strokePath(iconPen);
        buttonGraphics.newPath();
        buttonGraphics.rectPath(frontLeft, frontTop, sheetSize, sheetSize);
        buttonGraphics.strokePath(iconPen);
    }

    /**
     * コピーしたあとのチェックマークを描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCheckIcon(buttonGraphics, buttonSize) {
        var center = buttonSize / 2;
        buttonGraphics.newPath();
        buttonGraphics.moveTo(center - 4.5, center);
        buttonGraphics.lineTo(center - 1.5, center + 3);
        buttonGraphics.lineTo(center + 4.5, center - 3.5);
        buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1.5));
    }

    /**
     * 文字列をクリップボードにコピーする。
     * InDesign は AppleScript（Windows は VBScript）の doScript で書き込む。
     * Illustrator は文字列を直接送る API が無いので、一時レイヤーのテキストフレームを選んでコピーする（プレーンテキストも入る）。
     * Illustrator の常駐パレットからは DOM に触れないので、同じ処理を BridgeTalk でメインエンジンに渡す
     * @param {string} copyText - コピーする文字列
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {boolean} コピーの操作まで進めたら true
     */
    function copyVersionInfoText(copyText, isFromPalette) {
        try {
            if (VERSION_INFO_IS_INDESIGN) {
                if (File.fs === "Windows") {
                    app.doScript('CreateObject("htmlfile").parentWindow.clipboardData.setData "text", "' + copyText.replace(/"/g, '""') + '"', ScriptLanguage.VISUAL_BASIC);
                } else {
                    app.doScript('set the clipboard to "' + copyText.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"', ScriptLanguage.APPLESCRIPT_LANGUAGE);
                }
                return true;
            }
            var copyScript = buildVersionInfoCopyScript(copyText);
            if (!isFromPalette) return eval(copyScript) === true;
            var copyMessage = new BridgeTalk();
            copyMessage.target = BridgeTalk.appSpecifier; /* 複数バージョンを起動していても、今のバージョンに送る / the running version, even with several installed */
            copyMessage.body = copyScript;
            copyMessage.send();
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * Illustrator でコピーする処理を文字列で組み立てる（BridgeTalk でも送れるよう、関数の toString() は使わない）。
     * 一時レイヤーにテキストフレームを作り、フレームを選んでメニューコマンドでコピーして、レイヤーごと消す。
     * 文字として選択（textRange.select()）してコピーすると「オブジェクトをコピーできません」の警告が出るので使わない。
     * 選択と作業レイヤーは元に戻す。ドキュメントが無ければ一時ドキュメントを作って閉じる
     * @param {string} copyText - コピーする文字列
     * @returns {string} 実行すると true / false を返すスクリプト
     */
    function buildVersionInfoCopyScript(copyText) {
        return [
            "(function (copyText) {",
            "    var usingTempDoc = (app.documents.length === 0);",
            "    var targetDoc = usingTempDoc ? app.documents.add() : app.activeDocument;",
            "    var savedSelection = (!usingTempDoc && targetDoc.selection instanceof Array) ? targetDoc.selection : [];",
            "    var savedActiveLayer = usingTempDoc ? null : targetDoc.activeLayer;",
            "    var tempLayer = null;",
            "    var hasCopied = false;",
            "    try {",
            "        tempLayer = targetDoc.layers.add();",
            "        var tempFrame = tempLayer.textFrames.add();",
            "        tempFrame.contents = copyText;",
            "        app.redraw();",
            "        app.executeMenuCommand('deselectall');",
            "        tempFrame.selected = true;",
            "        app.redraw();",
            "        app.executeMenuCommand('copy');",
            "        app.redraw();",
            "        hasCopied = true;",
            "    } catch (e) { }",
            "    if (tempLayer) { try { tempLayer.remove(); } catch (e) { } }",
            "    if (usingTempDoc) { targetDoc.close(SaveOptions.DONOTSAVECHANGES); return hasCopied; }",
            "    try { targetDoc.activeLayer = savedActiveLayer; } catch (e) { }",
            "    targetDoc.selection = null;",
            "    for (var i = 0; i < savedSelection.length; i++) { try { savedSelection[i].selected = true; } catch (e) { } }",
            "    app.redraw();",
            "    return hasCopied;",
            "})(" + quoteVersionInfoString(copyText) + ");"
        ].join(" "); /* 行コメントを含まないので1行につなぐ（BridgeTalk の転送で改行が化けない）/ joined on one line; there are no line comments */
    }

    /**
     * 文字列を、スクリプトに埋め込める二重引用符のリテラルにする
     * @param {string} sourceText - 文字列
     * @returns {string} "…" の形のリテラル
     */
    function quoteVersionInfoString(sourceText) {
        return '"' + String(sourceText).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r/g, "\\r").replace(/\n/g, "\\n") + '"';
    }

    /**
     * ファイルを Finder で表示する。RevealInFinder.app があればファイルを選択した状態で表示し、無ければ入っているフォルダーを開く
     * @param {File} scriptFile - 表示するファイル
     * @returns {void}
     */
    function revealVersionInfoScriptFile(scriptFile) {
        /* .app は実体がフォルダーなので File.exists が false になる環境がある。Folder でも確かめる / An .app is a folder, so check with Folder too */
        var hasRevealApp = (File.fs === "Macintosh")
            && (new Folder(VERSION_INFO_REVEAL_APP_PATH).exists || new File(VERSION_INFO_REVEAL_APP_PATH).exists);
        /* アプリは一時ファイルからパスを読むので、書き出してから起動する / The app reads the path from a temp file */
        if (hasRevealApp && writeVersionInfoRevealPath(scriptFile.fsName) && new File(VERSION_INFO_REVEAL_APP_PATH).execute()) return;
        /* アプリが無い・起動できないときは、何も起きないままにせずフォルダーを開く / Otherwise open the enclosing folder */
        scriptFile.parent.execute();
    }

    /**
     * RevealInFinder.app に渡すパスを一時ファイルへ書き出す（UTF-8、改行なし）
     * @param {string} targetPath - 絶対パス
     * @returns {boolean} 書き出せたら true
     */
    function writeVersionInfoRevealPath(targetPath) {
        var pathFile = new File(VERSION_INFO_REVEAL_PATH_FILE);
        pathFile.encoding = "UTF-8";
        pathFile.lineFeed = "Unix";
        if (!pathFile.open("w")) return false;
        var hasWritten = pathFile.write(targetPath);
        pathFile.close();
        return hasWritten;
    }

    /**
     * URL を既定のブラウザーで開く（一時的なショートカットファイルを作って開く）
     * @param {string} targetUrl - 開く URL
     * @returns {void}
     */
    function openVersionInfoUrl(targetUrl) {
        var isMac = (File.fs === "Macintosh");
        var shortcutFile = new File(Folder.temp.fsName + "/VersionInfoLink" + (isMac ? ".webloc" : ".url"));
        shortcutFile.encoding = "UTF-8";
        shortcutFile.open("w");
        if (isMac) {
            var escapedUrl = targetUrl.replace(/&/g, "&amp;").replace(/</g, "&lt;");
            shortcutFile.write('<?xml version="1.0" encoding="UTF-8"?>\n'
                + '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n'
                + '<plist version="1.0"><dict><key>URL</key><string>' + escapedUrl + '</string></dict></plist>\n');
        } else {
            shortcutFile.write("[InternetShortcut]\r\nURL=" + targetUrl + "\r\n");
        }
        shortcutFile.close();
        shortcutFile.execute();
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "グラデーションストップの編集", en: "Edit Gradient Stops" },
            separateTitle: { ja: "セパレートグラデーション", en: "Segmented Gradient" }
        },
        panel: {
            untitledGradient: { ja: "グラデーション %1", en: "Gradient %1" },
            boundaries: { ja: "境界の位置", en: "Boundary positions" }
        },
        fieldLabel: {
            boundary: { ja: "境界 %1", en: "Boundary %1" }
        },
        column: {
            color: { ja: "カラー", en: "Color" },
            location: { ja: "位置", en: "Location" },
            opacity: { ja: "不透明度", en: "Opacity" }
        },
        colorText: {
            none: { ja: "なし", en: "None" },
            unknown: { ja: "（不明）", en: "(Unknown)" }
        },
        message: {
            omittedGradients: { ja: "ほか %1 個のグラデーションは省略しました。", en: "%1 more gradients are not shown." },
            modeCreate: { ja: "%1 色から新しいグラデーションを作成", en: "Create a new gradient from %1 colors" },
            modeAdjust: { ja: "%1 件のグラデーション（%2 色）を調整", en: "Adjust %1 gradient(s) with %2 colors" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            separate: { ja: "セパレート...", en: "Segment..." },
            setDashes: { ja: "丸型の破線", en: "Round Dashes" }
        },
        tooltip: {
            colorChip: { ja: "クリックしてカラーを変更", en: "Click to change the color" },
            setDashes: {
                ja: "グラデーションの線に、破線（線分 1pt・間隔は空白）・丸型線端・［コーナーやパス先端に破線の先端を整列］を設定",
                en: "Sets dashes (1 pt dash, blank gap), round caps and Align Dashes to Corners and Path Ends on gradient strokes"
            },
            separate: {
                ja: "色が混ざらない、境界のくっきりしたグラデーションにして、境界の位置を％で指定",
                en: "Make a segmented gradient with hard edges and set the boundary positions in %"
            },
            boundary: {
                ja: "前の色から次の色に切り替わる位置（グラデーションの始点 0%〜終点 100%）",
                en: "Where one color switches to the next (0% at the start of the gradient, 100% at the end)"
            },
            swapUp: { ja: "上のストップとカラー・不透明度を入れ替え", en: "Swap color and opacity with the stop above" },
            swapDown: { ja: "下のストップとカラー・不透明度を入れ替え", en: "Swap color and opacity with the stop below" },
            duplicate: { ja: "このストップを複製", en: "Duplicate this stop" },
            removeStop: { ja: "このストップを削除", en: "Delete this stop" },
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
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select an object." },
            noGradient: { ja: "選択したオブジェクトにグラデーションが見つかりません。", en: "No gradient was found on the selected objects." },
            applyFailed: { ja: "%1 個の塗り・線に変更を適用できませんでした。", en: "Could not apply the changes to %1 fills or strokes." },
            singleColor: {
                ja: "グラデーションの色が1色しかないため、境界がありません。",
                en: "The gradient has only one color, so there is no boundary to move."
            },
            tooManyColors: { ja: "色は %1 色までです。", en: "Up to %1 colors are supported." },
            invalidBoundaries: {
                ja: "境界の位置は 0% より大きく 100% より小さい値を、上の欄から順に大きくなるように入力してください。",
                en: "Enter boundary positions greater than 0% and less than 100%, increasing from the top field down."
            },
            dashAlignFailed: {
                ja: "［コーナーやパス先端に破線の先端を整列］を設定できませんでした。破線と丸型線端だけ設定しました。",
                en: "Could not set Align Dashes to Corners and Path Ends. Only the dashes and round caps were set."
            },
            skipped: {
                ja: "色数の異なるグラデーション %1 件は変更していません。",
                en: "%1 gradient(s) with a different number of colors were left unchanged."
            },
            noDrawableLayer: {
                ja: "長方形を描けるレイヤー（ロック・非表示でないもの）がありません。",
                en: "There is no unlocked, visible layer to draw the rectangle on."
            }
        }
    };

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
    // 前回の境界 / Last boundaries
    // =========================================

    /* 新規作成で使った境界を色数ごとに覚える（$.global に置き、Illustrator の終了まで残す）
       Boundaries used for new gradients, kept per color count until Illustrator quits */
    var LAST_BOUNDARIES_KEY = "__" + SCRIPT_NAME + "_LastBoundaries";

    // =========================================
    // 破線の一時アクション / Dash temporary action
    // =========================================

    var DASH_ACTION_SET_NAME = "SwwwitchTempDashSet";
    var DASH_ACTION_NAME     = "SwwwitchTempDash";
    var ACTION_UNIT_POINT    = 592476268;  /* ポイント / point（parameter /unit） */

    /* 線の設定（ai_plugin_setStroke）のパラメーターキー。記録した .aia から採取 / Parameter keys of ai_plugin_setStroke, taken from a recorded .aia */
    var STROKE_KEY_WIDTH      = 2003072104; /* wdth: 線幅 / stroke width */
    var STROKE_KEY_CAP        = 1667330094; /* cap : 線端 / cap */
    var STROKE_KEY_MITER      = 1836344690; /* mitr: 比率 / miter limit */
    var STROKE_KEY_JOIN       = 1785686382; /* join: 角の形状 / join */
    var STROKE_KEY_DASH_COUNT = 1684825454; /* dlen: 破線の値の数 / number of dash values */
    var STROKE_KEY_DASH_1     = 1685284913; /* dsh1: 線分 / first dash */
    var STROKE_KEY_DASH_ADJUST = 1684104298; /* dadj: コーナーやパス先端に破線の先端を整列 / align dashes to corners and path ends */
    var STROKE_KEY_ARROW_1    = 1634231345; /* ahd1: 始点の矢印 / start arrowhead */
    var STROKE_KEY_ARROW_2    = 1634231346; /* ahd2: 終点の矢印 / end arrowhead */
    var STROKE_KEY_TIP_ALIGN  = 1634230636; /* ahal: 矢印の配置 / arrow tip alignment */
    var STROKE_KEY_ALIGN      = 1634494318; /* algn: 線の位置 / stroke alignment */

    // =========================================
    // メイン処理 / Main
    // =========================================

    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var activeDoc = app.activeDocument;
    var selectionSource = activeDoc.selection;
    /* 文字ツールで文字を選択しているときは TextRange が返る / Selecting characters returns a TextRange */
    var isTextSelection = (selectionSource && selectionSource.typename === "TextRange");
    if (!selectionSource || (!isTextSelection && selectionSource.length === 0)) {
        alert(getLabel("alert.noSelection"));
        return;
    }

    var gradientEntries = collectGradientEntries(selectionSource);
    if (gradientEntries.length > 0) {
        showGradientStopsDialog(gradientEntries);
        return;
    }

    /* グラデーションが無く、色が2つ以上あれば、その色からセパレートグラデーションを作る / Without gradients, build a segmented gradient from two or more colors */
    var selectionOrientation = isTextSelection ? "unknown" : detectSelectionOrientation(selectionSource);
    var solidColors = isTextSelection ? [] : collectSolidColors(selectionSource, selectionOrientation);
    if (solidColors.length < 2) {
        alert(getLabel("alert.noGradient"));
        return;
    }
    runCreateMode(activeDoc, solidColors, selectionOrientation);

    // =========================================
    // グラデーションの収集 / Collecting gradients
    // =========================================

    /**
     * 選択からグラデーションを集める（塗りと線。同じストップ構成のものは1つにまとめ、使っている塗り・線を targets に控える）
     * @param {*} source - doc.selection
     * @returns {Object[]} { name, stops, targets } の配列（stops は位置順、targets は { item, paintKey }）
     */
    function collectGradientEntries(source) {
        var targetItems = collectSelectionItems(source, {
            accept: function (item) {
                return item.typename === "PathItem" || item.typename === "TextFrame";
            },
            enterCompoundPaths: true /* 複合パスの塗り・線は中のパスから読む / read compound path paint from its paths */
        });
        var gradientEntries = [];
        var entryBySignature = {};
        var paintKeys = ["fillColor", "strokeColor"];
        for (var i = 0; i < targetItems.length; i++) {
            var paintSource = getPaintSource(targetItems[i]);
            if (!paintSource) continue;
            for (var j = 0; j < paintKeys.length; j++) {
                var paintColor;
                try { paintColor = paintSource[paintKeys[j]]; } catch (e) { continue; }
                if (!paintColor || paintColor.typename !== "GradientColor") continue;
                var gradientEntry = readGradientEntry(paintColor.gradient);
                if (!gradientEntry) continue;
                var signature = buildGradientSignature(gradientEntry);
                if (entryBySignature.hasOwnProperty(signature)) {
                    gradientEntry = entryBySignature[signature];
                } else {
                    entryBySignature[signature] = gradientEntry;
                    gradientEntries.push(gradientEntry);
                }
                gradientEntry.targets.push({ item: targetItems[i], paintKey: paintKeys[j] });
            }
        }
        return gradientEntries;
    }

    /**
     * 塗り・線を読み書きする相手を返す（パスはそのもの、テキストは文字の属性）
     * @param {PageItem} item - パスまたはテキストフレーム
     * @returns {Object|null} fillColor / strokeColor を持つオブジェクト（読めなければ null）
     */
    function getPaintSource(item) {
        if (item.typename !== "TextFrame") return item;
        try {
            return item.textRange.characterAttributes;
        } catch (e) {
            return null;
        }
    }

    /**
     * グラデーションのストップを読み、位置の順に並べる
     * @param {Gradient} gradient - グラデーション
     * @returns {Object|null} { name, stops, targets, isWritten }（読めなければ null）
     */
    function readGradientEntry(gradient) {
        var stops = [];
        try {
            var gradientStops = gradient.gradientStops;
            for (var i = 0; i < gradientStops.length; i++) {
                var gradientStop = gradientStops[i];
                insertStopByRampPoint(stops, {
                    color: gradientStop.color,
                    rampPoint: gradientStop.rampPoint,
                    opacity: gradientStop.opacity,
                    midPoint: gradientStop.midPoint,
                    isAdded: false,        /* ダイアログで複製したストップ / stop duplicated in the dialog */
                    isRemoved: false,      /* ダイアログで削除したストップ（元に戻すときのために残す） / stop removed in the dialog, kept for rolling back */
                    editedColor: null,     /* ダイアログで選んだカラー / color picked in the dialog */
                    editedRampPoint: null, /* ダイアログで入れた位置 / location typed in the dialog */
                    editedOpacity: null    /* ダイアログで入れた不透明度 / opacity typed in the dialog */
                });
            }
        } catch (e) {
            return null;
        }
        return { name: gradient.name || "", stops: stops, targets: [], isWritten: false /* プレビューで書き込み済みか / written by the preview */ };
    }

    /**
     * ストップを位置の順を保って配列に入れる（同じ位置は元の順）
     * @param {Object[]} stops - 位置順のストップ
     * @param {Object} stopEntry - 入れるストップ（rampPoint を持つ）
     * @returns {void}
     */
    function insertStopByRampPoint(stops, stopEntry) {
        var insertIndex = stops.length;
        while (insertIndex > 0 && stops[insertIndex - 1].rampPoint > stopEntry.rampPoint) insertIndex--;
        stops.splice(insertIndex, 0, stopEntry);
    }

    /**
     * 同じグラデーションを見分けるための文字列を作る（ストップの色・位置・不透明度・中間点）
     * @param {Object} gradientEntry - readGradientEntry() の戻り値
     * @returns {string} 比較用の文字列
     */
    function buildGradientSignature(gradientEntry) {
        var parts = [];
        for (var i = 0; i < gradientEntry.stops.length; i++) {
            var stopEntry = gradientEntry.stops[i];
            parts.push(describeColor(stopEntry.color) + "@" + stopEntry.rampPoint + "/" + stopEntry.opacity + "~" + stopEntry.midPoint);
        }
        return parts.join("|");
    }

    // =========================================
    // 変更の適用 / Applying the edits
    // =========================================

    /**
     * ストップに変更があるか調べる
     * @param {Object} gradientEntry - collectGradientEntries() の要素
     * @returns {boolean} 1つでも変更があれば true
     */
    function hasStopEdits(gradientEntry) {
        for (var i = 0; i < gradientEntry.stops.length; i++) {
            var stopEntry = gradientEntry.stops[i];
            if (stopEntry.isAdded || stopEntry.isRemoved || stopEntry.editedColor || stopEntry.editedRampPoint !== null || stopEntry.editedOpacity !== null) return true;
        }
        return false;
    }

    /**
     * 書き込むストップを位置の順に並べて返す
     * @param {Object} gradientEntry - collectGradientEntries() の要素
     * @param {boolean} useEdits - true なら変更後の値（複製を含め、削除を除く）、false なら元の値
     * @returns {Object[]} { color, rampPoint, opacity, midPoint } の配列（位置順）
     */
    function buildStopsToWrite(gradientEntry, useEdits) {
        var stopsToWrite = [];
        for (var i = 0; i < gradientEntry.stops.length; i++) {
            var stopEntry = gradientEntry.stops[i];
            if (useEdits ? stopEntry.isRemoved : stopEntry.isAdded) continue;
            var hasEditedRampPoint = useEdits && stopEntry.editedRampPoint !== null;
            var hasEditedOpacity = useEdits && stopEntry.editedOpacity !== null;
            insertStopByRampPoint(stopsToWrite, {
                color: (useEdits && stopEntry.editedColor) ? stopEntry.editedColor : stopEntry.color,
                rampPoint: hasEditedRampPoint ? stopEntry.editedRampPoint : stopEntry.rampPoint,
                opacity: hasEditedOpacity ? stopEntry.editedOpacity : stopEntry.opacity,
                midPoint: stopEntry.midPoint
            });
        }
        return stopsToWrite;
    }

    /**
     * 変更をオブジェクトに書き込む／元に戻す。変更のあるグラデーションは変更後の値を書き、
     * 前に書き込んだのに変更が無くなった（または showEdits が false の）グラデーションは元の値に戻す
     * @param {Object[]} gradientEntries - collectGradientEntries() の戻り値
     * @param {boolean} showEdits - 変更を書き込むなら true、すべて元に戻すなら false
     * @returns {number} 書き込めなかった塗り・線の数
     */
    function syncGradientEdits(gradientEntries, showEdits) {
        var failedCount = 0;
        for (var i = 0; i < gradientEntries.length; i++) {
            failedCount += syncGradientEntry(gradientEntries[i], showEdits);
        }
        app.redraw();
        return failedCount;
    }

    /**
     * グラデーション1つ分の変更を書き込む／元に戻す（再描画はしない）
     * @param {Object} gradientEntry - collectGradientEntries() の要素
     * @param {boolean} showEdits - 変更を書き込むなら true、元に戻すなら false
     * @returns {number} 書き込めなかった塗り・線の数
     */
    function syncGradientEntry(gradientEntry, showEdits) {
        var failedCount = 0;
        if (showEdits && hasStopEdits(gradientEntry)) {
            failedCount = writeGradientEntry(gradientEntry, buildStopsToWrite(gradientEntry, true));
            gradientEntry.isWritten = true;
        } else if (gradientEntry.isWritten) {
            failedCount = writeGradientEntry(gradientEntry, buildStopsToWrite(gradientEntry, false));
            gradientEntry.isWritten = false;
        }
        return failedCount;
    }

    /**
     * ストップを、そのグラデーションを使っている塗り・線すべてに書き込む。
     * ストップはすべて位置の順に書き直すので、同じグラデーションを何度書いても結果は同じ
     * @param {Object} gradientEntry - collectGradientEntries() の要素
     * @param {Object[]} stopsToWrite - buildStopsToWrite() の戻り値
     * @returns {number} 書き込めなかった塗り・線の数
     */
    function writeGradientEntry(gradientEntry, stopsToWrite) {
        var failedCount = 0;
        for (var i = 0; i < gradientEntry.targets.length; i++) {
            var paintTarget = gradientEntry.targets[i];
            /* 書き込みが拒まれても（ロック中など）残りの塗り・線は続ける / keep going when a write is refused, e.g. on locked items */
            try {
                var paintSource = getPaintSource(paintTarget.item);
                var gradientColor = paintSource[paintTarget.paintKey];
                if (gradientColor.typename !== "GradientColor") {
                    failedCount++;
                    continue;
                }
                writeGradientStops(gradientColor.gradient.gradientStops, stopsToWrite);
                /* 書き直した GradientColor を入れ直して、オブジェクトに反映させる / reassign so the object picks up the change */
                paintSource[paintTarget.paintKey] = gradientColor;
            } catch (e) {
                $.writeln("writeGradientEntry: " + e);
                failedCount++;
            }
        }
        return failedCount;
    }

    /**
     * ストップに変更後の値を書き込む（数をそろえ、位置を隣を追い越さない順で動かしてから、位置の順に色・不透明度・中間点を書く）
     * @param {GradientStops} gradientStops - 書き込む先のストップ
     * @param {Object[]} stopsToWrite - buildStopsToWrite() の戻り値（位置の昇順）
     * @returns {void}
     */
    function writeGradientStops(gradientStops, stopsToWrite) {
        /* 複製・セパレートで増えた分を足し、元に戻すときは減らす / add stops for duplicates and segments, remove them when rolling back */
        while (gradientStops.length < stopsToWrite.length) gradientStops.add();
        while (gradientStops.length > stopsToWrite.length) gradientStops[gradientStops.length - 1].remove();
        var orderedStops = getStopsByPosition(gradientStops);
        moveStopsWithoutOverlap(orderedStops, stopsToWrite);
        var stopCount = orderedStops.length;
        for (var i = 0; i < stopCount; i++) {
            orderedStops[i].color = stopsToWrite[i].color;
            orderedStops[i].opacity = stopsToWrite[i].opacity;
            /* 最後のストップの中間点は使われないので、書けなくても続ける / the last stop's midpoint is unused */
            try { orderedStops[i].midPoint = stopsToWrite[i].midPoint; } catch (e) { }
        }
    }

    /**
     * グラデーションのストップを今の位置の昇順に並べた配列を返す（同じ位置は元の順のまま）
     * @param {GradientStops} gradientStops - グラデーションのストップ
     * @returns {GradientStop[]} 位置の昇順のストップ
     */
    function getStopsByPosition(gradientStops) {
        var orderedStops = [];
        for (var i = 0; i < gradientStops.length; i++) {
            var gradientStop = gradientStops[i];
            var insertAt = orderedStops.length;
            while (insertAt > 0 && orderedStops[insertAt - 1].rampPoint > gradientStop.rampPoint) insertAt--;
            orderedStops.splice(insertAt, 0, gradientStop);
        }
        return orderedStops;
    }

    /**
     * ストップを、隣のストップを追い越さない順で目標の位置へ動かす。
     * Illustrator は前後のストップを越える位置を入れると例外になるので、
     * 始点側へ動くものは始点側から、終点側へ動くものは終点側から順に動かす
     * @param {GradientStop[]} orderedStops - 位置の昇順のストップ
     * @param {Object[]} targetStops - 目標の {rampPoint} の配列（位置の昇順、orderedStops と同じ数）
     * @returns {void}
     */
    function moveStopsWithoutOverlap(orderedStops, targetStops) {
        var i;
        for (i = 0; i < orderedStops.length; i++) {
            if (targetStops[i].rampPoint < orderedStops[i].rampPoint) orderedStops[i].rampPoint = targetStops[i].rampPoint;
        }
        for (i = orderedStops.length - 1; i >= 0; i--) {
            if (targetStops[i].rampPoint > orderedStops[i].rampPoint) orderedStops[i].rampPoint = targetStops[i].rampPoint;
        }
    }

    // =========================================
    // カラーの表示 / Color display
    // =========================================

    /**
     * 数値を小数第1位で丸めた文字列にする（.0 は付けない）
     * @param {number} value - 数値
     * @returns {string} 表示用の文字列
     */
    function formatNumber(value) {
        return String(Math.round(value * 10) / 10);
    }

    /**
     * カラーの値を文字列にする（例「C0 M50 Y100 K0」「R255 G0 B0」「K50」、スウォッチは名前と濃度）
     * @param {Color} color - カラー
     * @returns {string} 表示用の文字列
     */
    function describeColor(color) {
        try {
            switch (color.typename) {
                case "CMYKColor":
                    return "C" + formatNumber(color.cyan) + " M" + formatNumber(color.magenta) +
                        " Y" + formatNumber(color.yellow) + " K" + formatNumber(color.black);
                case "RGBColor":
                    return "R" + formatNumber(color.red) + " G" + formatNumber(color.green) + " B" + formatNumber(color.blue);
                case "GrayColor":
                    return "K" + formatNumber(color.gray);
                case "LabColor":
                    return "L" + formatNumber(color.l) + " a" + formatNumber(color.a) + " b" + formatNumber(color.b);
                case "SpotColor":
                    /* グローバルプロセスカラーも SpotColor で返る / global process colors also come back as SpotColor */
                    return color.spot.name + (color.tint < 100 ? " " + formatNumber(color.tint) + "%" : "");
                case "NoColor":
                    return getLabel("colorText.none");
            }
        } catch (e) { /* 値を読めない / values are unreadable */ }
        return getLabel("colorText.unknown");
    }

    /**
     * カラーを画面表示用の RGB（0〜255）にする
     * @param {Color} color - カラー
     * @returns {number[]|null} [R, G, B]（変換できなければ null）
     */
    function toScreenRgb(color) {
        try {
            switch (color.typename) {
                case "RGBColor":
                    return [color.red, color.green, color.blue];
                case "CMYKColor":
                    return app.convertSampleColor(ImageColorSpace.CMYK, [color.cyan, color.magenta, color.yellow, color.black],
                        ImageColorSpace.RGB, ColorConvertPurpose.defaultpurpose);
                case "GrayColor":
                    return app.convertSampleColor(ImageColorSpace.GrayScale, [color.gray],
                        ImageColorSpace.RGB, ColorConvertPurpose.defaultpurpose);
                case "LabColor":
                    return app.convertSampleColor(ImageColorSpace.LAB, [color.l, color.a, color.b],
                        ImageColorSpace.RGB, ColorConvertPurpose.defaultpurpose);
                case "SpotColor":
                    var baseRgb = toScreenRgb(color.spot.color);
                    if (!baseRgb) return null;
                    /* 濃度の分だけ白に寄せる / lighten toward white by the tint */
                    var tintRatio = color.tint / 100;
                    return [
                        255 - (255 - baseRgb[0]) * tintRatio,
                        255 - (255 - baseRgb[1]) * tintRatio,
                        255 - (255 - baseRgb[2]) * tintRatio
                    ];
            }
        } catch (e) { /* 変換できない / cannot convert */ }
        return null;
    }

    // =========================================
    // セパレート / Segmented gradients
    // =========================================

    /**
     * 色を等分する境界の位置を返す（小数第1位まで）
     * @param {number} colorCount - 色数
     * @returns {number[]} 境界の位置（%、colorCount - 1 個）
     */
    function getEvenBoundaries(colorCount) {
        var boundaries = [];
        for (var i = 1; i < colorCount; i++) boundaries.push(Math.round(1000 * i / colorCount) / 10);
        return boundaries;
    }

    /**
     * 境界の位置が使えるか調べる（0〜100% の内側で、前後のストップが重ならずに昇順）
     * @param {number[]} boundaries - 境界の位置（%）
     * @returns {boolean} 使えるなら true
     */
    function areBoundariesValid(boundaries) {
        var previousEdge = 0;
        for (var i = 0; i < boundaries.length; i++) {
            var boundary = boundaries[i];
            if (typeof boundary !== "number" || isNaN(boundary)) return false;
            if (boundary - BOUNDARY_OFFSET <= previousEdge) return false;
            previousEdge = boundary + BOUNDARY_OFFSET;
        }
        return previousEdge < 100;
    }

    /**
     * 色の帯の並びと境界の位置から、セパレートグラデーションのストップ列を作る。
     * 両端に1つずつ、境界ごとに前の色と次の色を ±BOUNDARY_OFFSET の位置に置く
     * @param {Object[]} bands - {color, opacity} の配列（始点側から順）
     * @param {number[]} boundaries - 境界の位置（%、bands.length - 1 個、昇順）
     * @returns {Object[]} {color, rampPoint, opacity, midPoint} の配列（位置の昇順）
     */
    function buildSeparateStops(bands, boundaries) {
        var stops = [{ color: bands[0].color, rampPoint: 0, opacity: bands[0].opacity, midPoint: 50 }];
        for (var i = 0; i < boundaries.length; i++) {
            stops.push({ color: bands[i].color, rampPoint: boundaries[i] - BOUNDARY_OFFSET, opacity: bands[i].opacity, midPoint: 50 });
            stops.push({ color: bands[i + 1].color, rampPoint: boundaries[i] + BOUNDARY_OFFSET, opacity: bands[i + 1].opacity, midPoint: 50 });
        }
        var lastBand = bands[bands.length - 1];
        stops.push({ color: lastBand.color, rampPoint: 100, opacity: lastBand.opacity, midPoint: 50 });
        return stops;
    }

    /**
     * ストップ列を、同じ色・不透明度が続く「色の帯」にまとめる。
     * なめらかなグラデーションはストップ1つが1帯、セパレートは同じ色の2つのストップが1帯になる
     * @param {Object[]} stops - buildStopsToWrite() の戻り値（位置の昇順）
     * @returns {Object[]} {color, opacity, start, end} の配列（始点側から順）
     */
    function extractColorBands(stops) {
        var bands = [];
        var lastKey = null;
        for (var i = 0; i < stops.length; i++) {
            var bandKey = describeColor(stops[i].color) + "@" + stops[i].opacity;
            if (bandKey === lastKey) {
                bands[bands.length - 1].end = stops[i].rampPoint;
                continue;
            }
            lastKey = bandKey;
            bands.push({ color: stops[i].color, opacity: stops[i].opacity, start: stops[i].rampPoint, end: stops[i].rampPoint });
        }
        return bands;
    }

    /**
     * 境界の欄の初期値を返す。すでにセパレート（色の切り替わりがすべて SEPARATE_EDGE_MAX_GAP 以内）なら
     * 今の境界（前の帯の終わりと次の帯の始まりの中間、小数第2位まで）、なめらかなグラデーションなら等分
     * @param {Object[]} bands - extractColorBands() の戻り値
     * @returns {number[]} 境界の位置（%）
     */
    function getInitialBoundaries(bands) {
        var boundaries = [];
        for (var i = 0; i < bands.length - 1; i++) {
            if (bands[i + 1].start - bands[i].end > SEPARATE_EDGE_MAX_GAP) return getEvenBoundaries(bands.length);
            boundaries.push(Math.round((bands[i].end + bands[i + 1].start) * 50) / 100);
        }
        return boundaries;
    }

    /**
     * ダイアログで足すストップを作る（複製・セパレート用）
     * @param {Color} color - カラー
     * @param {number} rampPoint - 位置（0〜100）
     * @param {number} opacity - 不透明度（0〜100）
     * @param {number} midPoint - 中間点（13〜87）
     * @returns {Object} readGradientEntry() のストップと同じ形（isAdded は true）
     */
    function createAddedStop(color, rampPoint, opacity, midPoint) {
        return {
            color: color,
            rampPoint: rampPoint,
            opacity: opacity,
            midPoint: midPoint,
            isAdded: true,
            isRemoved: false,
            editedColor: null,
            editedRampPoint: null,
            editedOpacity: null
        };
    }

    /**
     * グラデーションのストップを、セパレートのストップに置き換える
     * （元からあるストップは削除の印を付けて残し、ダイアログで足したストップは消す）
     * @param {Object} gradientEntry - collectGradientEntries() の要素
     * @param {Object[]} bands - extractColorBands() の戻り値
     * @param {number[]} boundaries - 境界の位置（%）
     * @returns {void}
     */
    function replaceWithSeparateStops(gradientEntry, bands, boundaries) {
        var newStops = [];
        for (var i = 0; i < gradientEntry.stops.length; i++) {
            var stopEntry = gradientEntry.stops[i];
            if (stopEntry.isAdded) continue;
            stopEntry.isRemoved = true;
            newStops.push(stopEntry);
        }
        var separateStops = buildSeparateStops(bands, boundaries);
        for (var j = 0; j < separateStops.length; j++) {
            var separateStop = separateStops[j];
            newStops.push(createAddedStop(separateStop.color, separateStop.rampPoint, separateStop.opacity, separateStop.midPoint));
        }
        gradientEntry.stops = newStops;
    }

    /**
     * ストップの並びと削除の印を控える（セパレートのダイアログを取り消したとき用）
     * @param {Object} gradientEntry - collectGradientEntries() の要素
     * @returns {Object[]} {stopEntry, isRemoved} の配列
     */
    function saveStopList(gradientEntry) {
        var savedStops = [];
        for (var i = 0; i < gradientEntry.stops.length; i++) {
            savedStops.push({ stopEntry: gradientEntry.stops[i], isRemoved: gradientEntry.stops[i].isRemoved });
        }
        return savedStops;
    }

    /**
     * saveStopList() で控えた並びと削除の印に戻す
     * @param {Object} gradientEntry - collectGradientEntries() の要素
     * @param {Object[]} savedStops - saveStopList() の戻り値
     * @returns {void}
     */
    function restoreStopList(gradientEntry, savedStops) {
        var stops = [];
        for (var i = 0; i < savedStops.length; i++) {
            savedStops[i].stopEntry.isRemoved = savedStops[i].isRemoved;
            stops.push(savedStops[i].stopEntry);
        }
        gradientEntry.stops = stops;
    }

    /**
     * 一覧のグラデーションをセパレートにする。境界のダイアログを開き、［OK］で行をセパレートのストップに置き換える
     * （書き込みは親のダイアログの［OK］で。親の［キャンセル］では元に戻る）。
     * 対象は、2色以上ある最初のグラデーションと色数がそろうものだけ。
     * 対象に線のグラデーションがあれば、境界のダイアログに［丸型の破線］を置く（境界のダイアログの［キャンセル］で、そこで設定した分は戻す）
     * @param {Object[]} gradientPanels - addGradientPanel() の戻り値の配列
     * @param {Object} panelCallbacks - showGradientStopsDialog() の onRowsRebuilt / dashOriginals（［丸型の破線］を押す前の線の設定）を持つ
     * @returns {void}
     */
    function separateGradientPanels(gradientPanels, panelCallbacks) {
        /* 入力中の値を控えてから、パネルごとに色の帯を数える / capture typed values, then count the bands per panel */
        var panelBands = [];
        var colorCount = 0;
        for (var i = 0; i < gradientPanels.length; i++) {
            for (var j = 0; j < gradientPanels[i].rows.length; j++) readStopRowEdits(gradientPanels[i].rows[j]);
            panelBands.push(extractColorBands(buildStopsToWrite(gradientPanels[i].gradientEntry, true)));
            if (colorCount === 0 && panelBands[i].length >= 2) colorCount = panelBands[i].length;
        }
        if (colorCount === 0) {
            alert(getLabel("alert.singleColor"));
            return;
        }
        if (colorCount > MAX_SEPARATE_COLORS) {
            alert(getLabel("alert.tooManyColors", [MAX_SEPARATE_COLORS]));
            return;
        }

        /* 色数がそろうパネルだけを対象にし、取り消し用に並びを控える / only panels with the same color count, saved for rolling back */
        var separateTargets = [];
        for (var k = 0; k < gradientPanels.length; k++) {
            if (panelBands[k].length !== colorCount) continue;
            separateTargets.push({
                panelControls: gradientPanels[k],
                bands: panelBands[k],
                savedStops: saveStopList(gradientPanels[k].gradientEntry)
            });
        }

        /**
         * 対象を控えた並びに戻し、境界があればセパレートに置き換えて書き込む
         * @param {number[]|null} boundaries - 境界の位置（%。null なら元の並びのまま）
         * @returns {void}
         */
        function applyToTargets(boundaries) {
            for (var m = 0; m < separateTargets.length; m++) {
                var gradientEntry = separateTargets[m].panelControls.gradientEntry;
                restoreStopList(gradientEntry, separateTargets[m].savedStops);
                if (boundaries) replaceWithSeparateStops(gradientEntry, separateTargets[m].bands, boundaries);
                syncGradientEntry(gradientEntry, true);
            }
            app.redraw();
        }

        var targetEntries = [];
        for (var t = 0; t < separateTargets.length; t++) targetEntries.push(separateTargets[t].panelControls.gradientEntry);
        var strokeItems = collectGradientStrokeItems(targetEntries);
        var dashOriginals = panelCallbacks.dashOriginals;
        var dashOriginalCount = dashOriginals.length; /* このダイアログで初めて設定した線を見分ける / tells which strokes were first set in this dialog */
        var dashButton = null;
        if (strokeItems.length > 0) {
            dashButton = {
                label: getLabel("button.setDashes"),
                helpTip: getLabel("tooltip.setDashes"),
                onClick: function () {
                    addStrokeSnapshots(dashOriginals, strokeItems);
                    if (!applyDashSettings(strokeItems)) alert(getLabel("alert.dashAlignFailed"));
                }
            };
        }

        var boundaries = showBoundaryDialog({
            modeText: getLabel("message.modeAdjust", [separateTargets.length, colorCount]),
            initialBoundaries: getInitialBoundaries(separateTargets[0].bands),
            hasFooter: false,
            onPreview: applyToTargets,
            extraButton: dashButton
        });
        applyToTargets(boundaries);
        if (!boundaries) {
            /* このダイアログで設定した破線と線端を戻す / roll back the dashes and caps set in this dialog */
            restoreStrokeSnapshots(dashOriginals.slice(dashOriginalCount), true);
            dashOriginals.length = dashOriginalCount;
            return;
        }

        for (var n = 0; n < separateTargets.length; n++) buildStopRows(separateTargets[n].panelControls);
        panelCallbacks.onRowsRebuilt();
        var skippedCount = gradientPanels.length - separateTargets.length;
        if (skippedCount > 0) alert(getLabel("alert.skipped", [skippedCount]));
    }

    // =========================================
    // 破線の設定 / Dash settings
    // =========================================

    /**
     * 線がグラデーションのパスを集める（同じパスは1回だけ。テキストは外す）
     * @param {Object[]} gradientEntries - collectGradientEntries() の戻り値
     * @returns {PathItem[]} パス
     */
    function collectGradientStrokeItems(gradientEntries) {
        var strokeItems = [];
        for (var i = 0; i < gradientEntries.length; i++) {
            var paintTargets = gradientEntries[i].targets;
            for (var j = 0; j < paintTargets.length; j++) {
                var strokeItem = paintTargets[j].item;
                if (paintTargets[j].paintKey !== "strokeColor" || strokeItem.typename !== "PathItem") continue;
                var isListed = false;
                for (var k = 0; k < strokeItems.length; k++) {
                    if (strokeItems[k] === strokeItem) isListed = true;
                }
                if (!isListed) strokeItems.push(strokeItem);
            }
        }
        return strokeItems;
    }

    /**
     * 線の設定を控える（アクションで変わる線幅・角の形状・比率と、［キャンセル］で戻す破線・線端）
     * @param {PathItem[]} strokeItems - collectGradientStrokeItems() の戻り値
     * @returns {Object[]} {item, strokeWidth, strokeJoin, strokeMiterLimit, strokeDashes, strokeCap} の配列
     */
    function saveStrokeSnapshots(strokeItems) {
        var strokeSnapshots = [];
        for (var i = 0; i < strokeItems.length; i++) {
            var strokeItem = strokeItems[i];
            strokeSnapshots.push({
                item: strokeItem,
                strokeWidth: strokeItem.strokeWidth,
                strokeJoin: strokeItem.strokeJoin,
                strokeMiterLimit: strokeItem.strokeMiterLimit,
                strokeDashes: strokeItem.strokeDashes,
                strokeCap: strokeItem.strokeCap
            });
        }
        return strokeSnapshots;
    }

    /**
     * まだ控えていないパスの線の設定を、控えの配列に足す（最初に［丸型の破線］を押す前の値を残す）
     * @param {Object[]} strokeSnapshots - 控えの配列（saveStrokeSnapshots() と同じ形）
     * @param {PathItem[]} strokeItems - パス
     * @returns {void}
     */
    function addStrokeSnapshots(strokeSnapshots, strokeItems) {
        var newItems = [];
        for (var i = 0; i < strokeItems.length; i++) {
            var isSaved = false;
            for (var j = 0; j < strokeSnapshots.length; j++) {
                if (strokeSnapshots[j].item === strokeItems[i]) isSaved = true;
            }
            if (!isSaved) newItems.push(strokeItems[i]);
        }
        var newSnapshots = saveStrokeSnapshots(newItems);
        for (var k = 0; k < newSnapshots.length; k++) strokeSnapshots.push(newSnapshots[k]);
    }

    /**
     * 控えた線の設定に戻す
     * @param {Object[]} strokeSnapshots - saveStrokeSnapshots() の戻り値
     * @param {boolean} includeDashes - 破線と線端も戻すなら true（線幅・角の形状・比率だけなら false）
     * @returns {void}
     */
    function restoreStrokeSnapshots(strokeSnapshots, includeDashes) {
        for (var i = 0; i < strokeSnapshots.length; i++) {
            var snapshot = strokeSnapshots[i];
            /* ロック中などで戻せなくても、残りは続ける / keep going when one item refuses */
            try {
                snapshot.item.strokeWidth = snapshot.strokeWidth;
                snapshot.item.strokeJoin = snapshot.strokeJoin;
                snapshot.item.strokeMiterLimit = snapshot.strokeMiterLimit;
                if (includeDashes) {
                    snapshot.item.strokeDashes = snapshot.strokeDashes;
                    snapshot.item.strokeCap = snapshot.strokeCap;
                }
            } catch (e) { $.writeln("restoreStrokeSnapshots: " + e); }
        }
        app.redraw();
    }

    /**
     * パスに、破線（STROKE_DASH_PATTERN）・丸型線端・［コーナーやパス先端に破線の先端を整列］を設定する。
     * 整列は DOM に無いので線の設定のアクションで入れ、アクションが既定値に戻す線幅・角の形状・比率は控えた値に戻す。
     * 破線と線端は DOM でも確かめて入れ直す
     * @param {PathItem[]} strokeItems - collectGradientStrokeItems() の戻り値
     * @returns {boolean} 整列まで設定できたら true
     */
    function applyDashSettings(strokeItems) {
        var strokeSnapshots = saveStrokeSnapshots(strokeItems); /* アクションの前の線幅・角の形状・比率 / weight, join and miter before the action */
        var doc = app.activeDocument;
        var savedSelection = doc.selection;
        var isAligned = false;
        try {
            doc.selection = strokeItems;
            isAligned = runTemporaryAction(buildDashActionSource(strokeSnapshots[0].strokeWidth), DASH_ACTION_SET_NAME, DASH_ACTION_NAME);
        } catch (e) {
            $.writeln("applyDashSettings: " + e);
        }
        restoreStrokeSnapshots(strokeSnapshots, false);
        for (var i = 0; i < strokeItems.length; i++) {
            try {
                if (String(strokeItems[i].strokeDashes) !== String(STROKE_DASH_PATTERN)) strokeItems[i].strokeDashes = STROKE_DASH_PATTERN;
                if (strokeItems[i].strokeCap !== StrokeCap.ROUNDENDCAP) strokeItems[i].strokeCap = StrokeCap.ROUNDENDCAP;
            } catch (e2) { $.writeln("applyDashSettings: " + e2); }
        }
        try {
            if (savedSelection instanceof Array) doc.selection = savedSelection;
            else if (savedSelection && savedSelection.select) savedSelection.select();
        } catch (e3) { /* 選択を戻せなくても設定は済んでいる / the settings are already applied */ }
        app.redraw();
        return isAligned;
    }

    /**
     * 破線・丸型線端・［コーナーやパス先端に破線の先端を整列］を設定する、線の設定（ai_plugin_setStroke）の一時アクションを作る。
     * 省いたパラメーターは既定値に戻るので、すべて書く（線幅・角の形状・比率はあとで DOM で戻す。矢印はなし、線の位置は中央になる）
     * @param {number} strokeWidth - 線幅（pt）
     * @returns {string} アクション定義のテキスト
     */
    function buildDashActionSource(strokeWidth) {
        var actionLines = ["/version 3"]
            .concat(buildActionNameLines("", DASH_ACTION_SET_NAME))
            .concat(["/isOpen 1", "/actionCount 1", "/action-1 {"])
            .concat(buildActionNameLines("\t", DASH_ACTION_NAME))
            .concat(["\t/keyIndex 0", "\t/colorIndex 0", "\t/isOpen 1", "\t/eventCount 1", "\t/event-1 {",
                "\t\t/useRulersIn1stQuadrant 0", "\t\t/internalName (ai_plugin_setStroke)"])
            .concat(buildActionNameLines("\t\t", "線を設定", "localizedName"))
            .concat(["\t\t/isOpen 0", "\t\t/isOn 1", "\t\t/hasDialog 0", "\t\t/parameterCount " + (10 + STROKE_DASH_PATTERN.length)]);
        var paramIndex = 0;

        /**
         * パラメーターを1つ足す
         * @param {number} paramKey - パラメーターキー
         * @param {string[]} bodyLines - /type から後の行（字下げなし）
         * @returns {void}
         */
        function addParam(paramKey, bodyLines) {
            paramIndex++;
            actionLines.push("\t\t/parameter-" + paramIndex + " {", "\t\t\t/key " + paramKey, "\t\t\t/showInPalette 4294967295");
            for (var i = 0; i < bodyLines.length; i++) actionLines.push("\t\t\t" + bodyLines[i]);
            actionLines.push("\t\t}");
        }

        /**
         * 列挙パラメーターの本文を作る
         * @param {string} displayName - 表示名
         * @param {number} enumValue - 値
         * @returns {string[]} 本文の行
         */
        function enumBody(displayName, enumValue) {
            return ["/type (enumerated)"].concat(buildActionNameLines("", displayName), ["/value " + enumValue]);
        }

        addParam(STROKE_KEY_WIDTH, ["/type (unit real)", "/value " + formatActionReal(strokeWidth), "/unit " + ACTION_UNIT_POINT]);
        addParam(STROKE_KEY_CAP, enumBody("丸型線端", 1));
        addParam(STROKE_KEY_MITER, ["/type (real)", "/value 10.0"]);
        addParam(STROKE_KEY_JOIN, enumBody("マイター結合", 0));
        addParam(STROKE_KEY_DASH_COUNT, ["/type (integer)", "/value " + STROKE_DASH_PATTERN.length]);
        for (var d = 0; d < STROKE_DASH_PATTERN.length; d++) {
            addParam(STROKE_KEY_DASH_1 + d, ["/type (unit real)", "/value " + formatActionReal(STROKE_DASH_PATTERN[d]), "/unit " + ACTION_UNIT_POINT]);
        }
        addParam(STROKE_KEY_DASH_ADJUST, ["/type (boolean)", "/value 1"]);
        /* 矢印は「[なし]」 / no arrowheads */
        addParam(STROKE_KEY_ARROW_1, ["/type (ustring)"].concat(buildActionNameLines("", "[なし]", "value")));
        addParam(STROKE_KEY_ARROW_2, ["/type (ustring)"].concat(buildActionNameLines("", "[なし]", "value")));
        addParam(STROKE_KEY_TIP_ALIGN, enumBody("パスの終点に配置", 0));
        addParam(STROKE_KEY_ALIGN, enumBody("中央", 0));
        actionLines.push("\t}", "}", "");
        return actionLines.join(String.fromCharCode(13));
    }

    /**
     * アクション定義の実数の書き方にする（整数でも「.0」を付ける）
     * @param {number} value - 数値
     * @returns {string} 「1.0」「0.5」の形の文字列
     */
    function formatActionReal(value) {
        var realText = String(Math.round(value * 10000) / 10000);
        return (realText.indexOf(".") < 0) ? realText + ".0" : realText;
    }

    // =========================================
    // 新規作成 / Creating gradients
    // =========================================

    /**
     * 選択から、ストップに使える単色（グラデーション・パターン・なし以外）を、
     * 長辺方向（横並びは左→右／縦並びは上→下）の順に重複を除いて集める
     * @param {*} source - doc.selection
     * @param {string} selectionOrientation - detectSelectionOrientation() の結果
     * @returns {Color[]} 重複を除いた色
     */
    function collectSolidColors(source, selectionOrientation) {
        var colorEntries = [];
        var leafItems = collectSelectionItems(source, {
            enterCompoundPaths: true,
            textRangeToFrame: false,
            accept: function (item) {
                return item.typename !== "GroupItem" && item.typename !== "CompoundPathItem";
            }
        });
        for (var i = 0; i < leafItems.length; i++) collectColorEntries(leafItems[i], colorEntries);

        /* 縦並びなら上→下、それ以外は左→右を優先。数件なので挿入ソート（比較関数つきの sort() は使わない）
           top-to-bottom for a vertical selection, otherwise left-to-right; insertion sort instead of sort() with a comparator */
        var isVertical = (selectionOrientation === "vertical");
        var orderedEntries = [];
        for (var j = 0; j < colorEntries.length; j++) {
            var insertAt = orderedEntries.length;
            while (insertAt > 0 && isColorEntryBefore(colorEntries[j], orderedEntries[insertAt - 1], isVertical)) insertAt--;
            orderedEntries.splice(insertAt, 0, colorEntries[j]);
        }

        var colors = [];
        var seenKeys = {};
        for (var k = 0; k < orderedEntries.length; k++) {
            var entryColor = orderedEntries[k].color;
            if (!entryColor) continue;
            var colorType = entryColor.typename;
            if (colorType === "NoColor" || colorType === "GradientColor" || colorType === "PatternColor") continue;
            var colorText = describeColor(entryColor);
            if (seenKeys.hasOwnProperty(colorText)) continue;
            seenKeys[colorText] = true;
            colors.push(entryColor);
        }
        return colors;
    }

    /**
     * 色の並び順で、entryA を entryB より前に置くかを返す
     * @param {{left: number, top: number}} entryA - 色の項目
     * @param {{left: number, top: number}} entryB - 色の項目
     * @param {boolean} isVertical - 縦並び（上→下を優先）なら true、それ以外は左→右を優先
     * @returns {boolean} entryA が先なら true（同じ位置なら false で元の順を保つ）
     */
    function isColorEntryBefore(entryA, entryB, isVertical) {
        if (isVertical) {
            if (entryA.top !== entryB.top) return entryA.top > entryB.top;
            return entryA.left < entryB.left;
        }
        if (entryA.left !== entryB.left) return entryA.left < entryB.left;
        return entryA.top > entryB.top;
    }

    /**
     * オブジェクト1件から塗り・線の色を、左上の位置付きで集める
     * @param {PageItem} item - 対象のオブジェクト（グループ・複合パスは展開済み）
     * @param {Object[]} outEntries - {left, top, color} を足す配列
     * @returns {void}
     */
    function collectColorEntries(item, outEntries) {
        try {
            var bounds = item.geometricBounds;
            var paintSource = getPaintSource(item);
            if (!paintSource) return;
            var isText = (item.typename === "TextFrame");
            /* テキストは filled / stroked を持たないので、色があれば使う / text has no filled / stroked flags */
            if (isText || item.filled) outEntries.push({ left: bounds[0], top: bounds[1], color: paintSource.fillColor });
            if (isText || item.stroked) outEntries.push({ left: bounds[0], top: bounds[1], color: paintSource.strokeColor });
        } catch (e) { /* 色や位置を読めないオブジェクトは外す / skip items whose color or bounds are unreadable */ }
    }

    /**
     * 選択オブジェクトが横並びか縦並びかを、各オブジェクトの中心の散らばりから判定する
     * @param {*} source - doc.selection
     * @returns {string} "horizontal" / "vertical" / "mixed" / "unknown"
     */
    function detectSelectionOrientation(source) {
        var selectedItems = normalizeSelectionItems(source);
        if (selectedItems.length < 2) return "unknown";
        var minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
        for (var i = 0; i < selectedItems.length; i++) {
            try {
                var bounds = selectedItems[i].geometricBounds;
                var centerX = (bounds[0] + bounds[2]) / 2;
                var centerY = (bounds[1] + bounds[3]) / 2;
                if (centerX < minX) minX = centerX;
                if (centerX > maxX) maxX = centerX;
                if (centerY < minY) minY = centerY;
                if (centerY > maxY) maxY = centerY;
            } catch (e) { /* 境界を読めないものは外す / skip items without bounds */ }
        }
        if (minX > maxX) return "unknown";
        var dx = maxX - minX;
        var dy = maxY - minY;
        if (dx > dy) return "horizontal";
        if (dy > dx) return "vertical";
        return "mixed";
    }

    /**
     * 前回同じ色数で使った境界を返す（無いか使えなければ等分）
     * @param {number} colorCount - 色数
     * @returns {number[]} 境界の位置（%）
     */
    function loadLastBoundaries(colorCount) {
        var savedSets = $.global[LAST_BOUNDARIES_KEY];
        var savedBoundaries = savedSets ? savedSets[colorCount] : null;
        if (savedBoundaries instanceof Array && savedBoundaries.length === colorCount - 1 && areBoundariesValid(savedBoundaries)) {
            return savedBoundaries.slice(0);
        }
        return getEvenBoundaries(colorCount);
    }

    /**
     * 境界を色数ごとに覚える
     * @param {number[]} boundaries - 境界の位置（%）
     * @returns {void}
     */
    function saveLastBoundaries(boundaries) {
        if (!$.global[LAST_BOUNDARIES_KEY]) $.global[LAST_BOUNDARIES_KEY] = {};
        $.global[LAST_BOUNDARIES_KEY][boundaries.length + 1] = boundaries.slice(0);
    }

    /**
     * 選択の色から新しいセパレートグラデーションを作る。境界のダイアログを開き、［OK］で長方形に塗って置く
     * @param {Document} doc - 対象ドキュメント
     * @param {Color[]} colors - collectSolidColors() の戻り値（2色以上）
     * @param {string} selectionOrientation - detectSelectionOrientation() の結果
     * @returns {void}
     */
    function runCreateMode(doc, colors, selectionOrientation) {
        if (colors.length > MAX_SEPARATE_COLORS) {
            alert(getLabel("alert.tooManyColors", [MAX_SEPARATE_COLORS]));
            return;
        }
        /* クリップグループはマスクで測る / clip groups are measured by their mask */
        var selectionBounds = getClipAwareUnionBounds(doc.selection, false);

        var boundaries = showBoundaryDialog({
            modeText: getLabel("message.modeCreate", [colors.length]),
            initialBoundaries: loadLastBoundaries(colors.length),
            hasFooter: true
        });
        if (!boundaries) return;

        saveLastBoundaries(boundaries);
        createSeparateGradientRect(doc, colors, boundaries, selectionBounds, selectionOrientation);
    }

    /**
     * 色と境界からセパレートグラデーションを作り、長方形に塗る（縦並びなら角度を 90° に）
     * @param {Document} doc - 対象ドキュメント
     * @param {Color[]} colors - 始点側から並べる色
     * @param {number[]} boundaries - 境界の位置（%）
     * @param {number[]|null} selectionBounds - 選択の [左, 上, 右, 下]
     * @param {string} selectionOrientation - detectSelectionOrientation() の結果
     * @returns {void}
     */
    function createSeparateGradientRect(doc, colors, boundaries, selectionBounds, selectionOrientation) {
        var targetLayer = getUnlockedVisibleLayer(doc);
        if (!targetLayer) {
            alert(getLabel("alert.noDrawableLayer"));
            return;
        }

        var bands = [];
        for (var i = 0; i < colors.length; i++) bands.push({ color: colors[i], opacity: 100 });

        var gradient = doc.gradients.add();
        gradient.type = GradientType.LINEAR;
        writeGradientStops(gradient.gradientStops, buildSeparateStops(bands, boundaries));
        gradient.name = uniqueName(NEW_GRADIENT_BASE_NAME, doc.gradients);

        var rectFrame = computeRectFrame(doc, selectionBounds, selectionOrientation);
        var gradientRect = targetLayer.pathItems.rectangle(rectFrame.top, rectFrame.left, rectFrame.width, rectFrame.height);
        doc.selection = null;
        gradientRect.selected = true;

        /* 縦並びなら上→下の向きにする（塗る前に実行） / top-to-bottom for a vertical selection, run before filling */
        if (selectionOrientation === "vertical") runGradientAngle90Action();

        gradientRect.stroked = false;
        gradientRect.filled = true;
        var gradientFill = new GradientColor();
        gradientFill.gradient = gradient;
        gradientRect.fillColor = gradientFill;
    }

    /**
     * 長方形の位置と大きさを決める。大きさは選択の外接サイズ、
     * 縦並びなら選択の右、それ以外は選択の下に、長方形1個分の間隔をあけて置く
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]|null} selectionBounds - 選択の [左, 上, 右, 下]（測れなければ null）
     * @param {string} selectionOrientation - detectSelectionOrientation() の結果
     * @returns {{left: number, top: number, width: number, height: number}} 長方形の左上と大きさ
     */
    function computeRectFrame(doc, selectionBounds, selectionOrientation) {
        if (!selectionBounds || selectionBounds[0] > selectionBounds[2] || selectionBounds[3] > selectionBounds[1]) {
            var viewCenter = doc.activeView.centerPoint;
            return {
                left: viewCenter[0] - DEFAULT_RECT_SIZE / 2,
                top: viewCenter[1] + DEFAULT_RECT_SIZE / 2,
                width: DEFAULT_RECT_SIZE,
                height: DEFAULT_RECT_SIZE
            };
        }
        var rectWidth = selectionBounds[2] - selectionBounds[0];
        var rectHeight = selectionBounds[1] - selectionBounds[3];
        if (!(rectWidth > 0)) rectWidth = DEFAULT_RECT_SIZE;
        if (!(rectHeight > 0)) rectHeight = DEFAULT_RECT_SIZE;

        if (selectionOrientation === "vertical") {
            return { left: selectionBounds[2] + rectWidth, top: selectionBounds[1], width: rectWidth, height: rectHeight };
        }
        return { left: selectionBounds[0], top: selectionBounds[3] - rectHeight, width: rectWidth, height: rectHeight };
    }

    /**
     * 描画できるレイヤー（ロックも非表示もされていないもの）を返す。アクティブレイヤーを優先する
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer|null} 描画できるレイヤー（無ければ null）
     */
    function getUnlockedVisibleLayer(doc) {
        var activeLayer = doc.activeLayer;
        if (activeLayer && !activeLayer.locked && activeLayer.visible) return activeLayer;
        for (var i = 0; i < doc.layers.length; i++) {
            if (!doc.layers[i].locked && doc.layers[i].visible) return doc.layers[i];
        }
        return null;
    }

    /**
     * 名前の衝突を避けた一意な名前を作る（"名前", "名前 1", "名前 2" …）
     * @param {string} baseName - 元の名前
     * @param {Object} namedCollection - 名前の重複を調べるコレクション（getByName を持つ）
     * @returns {string} 使われていない名前
     */
    function uniqueName(baseName, namedCollection) {
        var candidateName = baseName;
        for (var suffix = 1; hasItemNamed(namedCollection, candidateName); suffix++) {
            candidateName = baseName + " " + suffix;
        }
        return candidateName;
    }

    /**
     * コレクションに同じ名前の項目があるかを調べる
     * @param {Object} namedCollection - getByName を持つコレクション
     * @param {string} itemName - 調べる名前
     * @returns {boolean} あれば true
     */
    function hasItemNamed(namedCollection, itemName) {
        try {
            namedCollection.getByName(itemName); /* 見つからないと例外 / throws when not found */
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 選択中のオブジェクトのグラデーション角度を 90° にするアクションを実行する
     * @returns {void}
     */
    function runGradientAngle90Action() {
        var actionCode = [
            " /version 3",
            "/name [ 8",
            "\t6772616469656e74",
            "]",
            "/isOpen 1",
            "/actionCount 1",
            "/action-1 {",
            "\t/name [ 8",
            "\t\t3930646567726565",
            "\t]",
            "\t/keyIndex 0",
            "\t/colorIndex 0",
            "\t/isOpen 1",
            "\t/eventCount 1",
            "\t/event-1 {",
            "\t\t/useRulersIn1stQuadrant 0",
            "\t\t/internalName (ai_plugin_setGradient)",
            "\t\t/localizedName [ 30",
            "\t\t\te382b0e383a9e38387e383bce382b7e383a7e383b3e38292e8a8ade5ae9a",
            "\t\t]",
            "\t\t/isOpen 1",
            "\t\t/isOn 1",
            "\t\t/hasDialog 0",
            "\t\t/parameterCount 1",
            "\t\t/parameter-1 {",
            "\t\t\t/key 1634625388",
            "\t\t\t/showInPalette 4294967295",
            "\t\t\t/type (unit real)",
            "\t\t\t/value -90.0",
            "\t\t\t/unit 591490663",
            "\t\t}",
            "\t}",
            "}",
            ""
        ].join(String.fromCharCode(13));
        runTemporaryAction(actionCode, "gradient", "90degree");
    }

    // =========================================
    // ダイアログUI / Dialog UI
    // =========================================

    /**
     * グラデーションストップの一覧をダイアログボックスに表示し、変更を受け付ける。
     * 変更はその場でオブジェクトに書き込み（プレビューは常にオン）、［OK］で確定、［キャンセル］で元に戻す
     * @param {Object[]} gradientEntries - collectGradientEntries() の戻り値
     * @returns {void}
     */
    function showGradientStopsDialog(gradientEntries) {
        var stopsDialog = new Window("dialog", getLabel("dialog.title"));
        setupWindow(stopsDialog);

        var gradientPanels = [];
        var panelCallbacks = {
            /* 行が変わったら書き込み直す / rewrite whenever a row changes */
            onStopEdited: function () {
                syncDialogEdits(gradientEntries, gradientPanels, true);
            },
            /* 行を作り直したあと、ダイアログの大きさと見出しを合わせ直す / refit the dialog and headers after rebuilding rows */
            onRowsRebuilt: function () {
                stopsDialog.layout.layout(true);
                stopsDialog.layout.resize();
                alignAllHeaderLabels(gradientPanels);
            },
            /* セパレートのダイアログで［丸型の破線］を押す前の線の設定（［キャンセル］で戻す） / strokes before Round Dashes, restored on Cancel */
            dashOriginals: []
        };

        var listedCount = Math.min(gradientEntries.length, MAX_LISTED_GRADIENTS);
        for (var i = 0; i < listedCount; i++) {
            gradientPanels.push(addGradientPanel(stopsDialog, gradientEntries[i], i + 1, panelCallbacks));
        }
        if (gradientEntries.length > listedCount) {
            stopsDialog.add("statictext", undefined, getLabel("message.omittedGradients", [gradientEntries.length - listedCount]));
        }

        var buttonRow = addButtonRow(stopsDialog);
        var separateButton = buttonRow.leftGroup.add("button", undefined, getLabel("button.separate"));
        separateButton.helpTip = getLabel("tooltip.separate");
        separateButton.onClick = function () {
            separateGradientPanels(gradientPanels, panelCallbacks);
        };
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, "OK", { name: "ok" });

        alignRightOnlyButtonRow(buttonRow);
        /* レイアウトが決まってから見出しを入力欄に揃える（ボタン行の並べ直しより後） / align the headers once the layout is settled */
        stopsDialog.addEventListener("show", function () {
            alignAllHeaderLabels(gradientPanels);
        });
        addVersionFooter(stopsDialog);
        prepareDialogWindow(stopsDialog, SCRIPT_NAME);
        var isConfirmed = (stopsDialog.show() === 1);

        /* OK なら変更を書き込み、キャンセルならプレビューで書いた分を元に戻す / write on OK, roll back the preview on Cancel */
        var failedCount = syncDialogEdits(gradientEntries, gradientPanels, isConfirmed);
        /* セパレートのダイアログで設定した破線と線端を戻す / roll back the dashes and caps set in the segment dialog */
        if (!isConfirmed && panelCallbacks.dashOriginals.length > 0) restoreStrokeSnapshots(panelCallbacks.dashOriginals, true);
        if (isConfirmed && failedCount > 0) alert(getLabel("alert.applyFailed", [failedCount]));
    }

    /**
     * すべての行の入力を読んでから、変更を書き込む／元に戻す
     * @param {Object[]} gradientEntries - collectGradientEntries() の戻り値
     * @param {Object[]} gradientPanels - addGradientPanel() の戻り値の配列
     * @param {boolean} showEdits - 変更を書き込むなら true、元に戻すなら false
     * @returns {number} 書き込めなかった塗り・線の数
     */
    function syncDialogEdits(gradientEntries, gradientPanels, showEdits) {
        for (var i = 0; i < gradientPanels.length; i++) {
            for (var j = 0; j < gradientPanels[i].rows.length; j++) readStopRowEdits(gradientPanels[i].rows[j]);
        }
        return syncGradientEdits(gradientEntries, showEdits);
    }

    /**
     * すべてのパネルの見出しを入力欄に揃える
     * @param {Object[]} gradientPanels - addGradientPanel() の戻り値の配列
     * @returns {void}
     */
    function alignAllHeaderLabels(gradientPanels) {
        for (var i = 0; i < gradientPanels.length; i++) {
            alignHeaderLabels(gradientPanels[i].headerControls, gradientPanels[i].rows[0]);
        }
    }

    /**
     * グラデーション1つ分のパネル（見出しとストップの行）を足す
     * @param {Window} stopsDialog - 追加先のダイアログ
     * @param {Object} gradientEntry - collectGradientEntries() の要素
     * @param {number} gradientNumber - 名前の無いグラデーションに付ける番号（1始まり）
     * @param {Object} panelCallbacks - onStopEdited（値が変わったとき）/ onRowsRebuilt（行を作り直したとき）
     * @returns {{gradientPanel: Panel, gradientEntry: Object, headerControls: Object, rows: Object[], panelCallbacks: Object}} パネルの部品（rows は上から順、作り直すと入れ替わる）
     */
    function addGradientPanel(stopsDialog, gradientEntry, gradientNumber, panelCallbacks) {
        var panelTitle = gradientEntry.name || getLabel("panel.untitledGradient", [gradientNumber]);
        var gradientPanel = stopsDialog.add("panel", undefined, panelTitle);
        setupPanel(gradientPanel, STOP_ROW_SPACING);
        var panelControls = {
            gradientPanel: gradientPanel,
            gradientEntry: gradientEntry,
            headerControls: addStopHeaderRow(gradientPanel),
            rows: [],
            panelCallbacks: panelCallbacks
        };
        buildStopRows(panelControls);
        return panelControls;
    }

    /**
     * パネルのストップの行を、gradientEntry.stops の並びで作り直す（見出しの行は残し、削除したストップは出さない）
     * @param {Object} panelControls - addGradientPanel() の戻り値
     * @returns {void}
     */
    function buildStopRows(panelControls) {
        for (var i = 0; i < panelControls.rows.length; i++) {
            panelControls.gradientPanel.remove(panelControls.rows[i].rowGroup);
        }
        var stops = panelControls.gradientEntry.stops;
        var panelRows = [];
        for (var j = 0; j < stops.length; j++) {
            if (stops[j].isRemoved) continue;
            panelRows.push(addStopRow(panelControls.gradientPanel, stops[j], panelControls.panelCallbacks.onStopEdited));
        }
        panelControls.rows = panelRows;
        /* 入れ替え・複製は隣の行を参照するので、行がそろってから足す / added once all rows exist, since they reference neighbours */
        for (var k = 0; k < panelRows.length; k++) addRowButtons(panelControls, k);
    }

    /**
     * ストップを複製して、すぐ下に入れる（位置は次のストップとの中間、最後の行なら前のストップとの中間）
     * @param {Object} panelControls - addGradientPanel() の戻り値
     * @param {number} rowIndex - 複製する行の番号
     * @returns {void}
     */
    function duplicateStopRow(panelControls, rowIndex) {
        var panelRows = panelControls.rows;
        /* 作り直す前に、入力中の値を控える / capture typed values before the rows are rebuilt */
        for (var i = 0; i < panelRows.length; i++) readStopRowEdits(panelRows[i]);
        var sourceStop = panelRows[rowIndex].stopEntry;
        var neighbourIndex = (rowIndex < panelRows.length - 1) ? rowIndex + 1 : rowIndex - 1;
        var sourceLocation = getCurrentRampPoint(sourceStop);
        var neighbourLocation = getCurrentRampPoint(panelRows[neighbourIndex].stopEntry);
        /* 削除したストップも stops に残っているので、行の番号ではなく元のストップの場所の次に入れる / insert after the source stop, since removed stops stay in the list */
        var stops = panelControls.gradientEntry.stops;
        stops.splice(indexOfStop(stops, sourceStop) + 1, 0, createAddedStop(
            sourceStop.editedColor || sourceStop.color,
            Math.round((sourceLocation + neighbourLocation) / 2 * 100) / 100,
            (sourceStop.editedOpacity !== null) ? sourceStop.editedOpacity : sourceStop.opacity,
            50
        ));
        rebuildPanelRows(panelControls);
    }

    /**
     * ストップを削除する（複製したストップは消し、元からあるストップは削除の印を付ける）。ストップは2つより少なくできない
     * @param {Object} panelControls - addGradientPanel() の戻り値
     * @param {number} rowIndex - 削除する行の番号
     * @returns {void}
     */
    function removeStopRow(panelControls, rowIndex) {
        var panelRows = panelControls.rows;
        if (panelRows.length <= MIN_STOP_COUNT) return;
        /* 作り直す前に、入力中の値を控える / capture typed values before the rows are rebuilt */
        for (var i = 0; i < panelRows.length; i++) readStopRowEdits(panelRows[i]);
        var targetStop = panelRows[rowIndex].stopEntry;
        if (targetStop.isAdded) {
            var stops = panelControls.gradientEntry.stops;
            stops.splice(indexOfStop(stops, targetStop), 1);
        } else {
            targetStop.isRemoved = true;
        }
        rebuildPanelRows(panelControls);
    }

    /**
     * 行を作り直し、ダイアログの大きさを合わせ直して、プレビューを更新する
     * @param {Object} panelControls - addGradientPanel() の戻り値
     * @returns {void}
     */
    function rebuildPanelRows(panelControls) {
        buildStopRows(panelControls);
        panelControls.panelCallbacks.onRowsRebuilt();
        panelControls.panelCallbacks.onStopEdited();
    }

    /**
     * ストップが stops の何番目にあるかを返す
     * @param {Object[]} stops - gradientEntry.stops
     * @param {Object} stopEntry - 探すストップ
     * @returns {number} 番号（無ければ -1）
     */
    function indexOfStop(stops, stopEntry) {
        for (var i = 0; i < stops.length; i++) {
            if (stops[i] === stopEntry) return i;
        }
        return -1;
    }

    /**
     * ストップの今の位置を返す（入力で変えていればその値）
     * @param {Object} stopEntry - readGradientEntry() のストップ
     * @returns {number} 位置（0〜100）
     */
    function getCurrentRampPoint(stopEntry) {
        return (stopEntry.editedRampPoint !== null) ? stopEntry.editedRampPoint : stopEntry.rampPoint;
    }

    /**
     * ストップの行と同じ幅の欄を持つ行を作る
     * @param {Panel} parent - 追加先のパネル
     * @returns {Group} 行
     */
    function addStopRowGroup(parent) {
        var stopRow = parent.add("group");
        setupRow(stopRow, "left", STOP_CELL_SPACING);
        return stopRow;
    }

    /**
     * 決まった幅の文字欄を足す
     * @param {Group} stopRow - 行
     * @param {string} text - 文字
     * @param {number} width - 幅（px）
     * @param {boolean} [alignRight] - 右揃えにするなら true
     * @returns {StaticText} 文字欄
     */
    function addStopCell(stopRow, text, width, alignRight) {
        var cellText = stopRow.add("statictext", undefined, text);
        cellText.preferredSize.width = width;
        if (alignRight) cellText.justify = "right";
        return cellText;
    }

    /**
     * 見出しの行（カラー・位置・不透明度）を足す。位置・不透明度の見出しは、表示したときに alignHeaderLabels() で入力欄に揃える
     * @param {Panel} parent - 追加先のパネル
     * @returns {{headerRow: Group, locationLabel: StaticText, opacityLabel: StaticText}} 行と見出し
     */
    function addStopHeaderRow(parent) {
        var headerRow = addStopRowGroup(parent);
        /* 見本の分を空ける / leave room for the color chip */
        var chipSpacer = headerRow.add("group");
        chipSpacer.preferredSize = [COLOR_CHIP_SIZE, 1];
        addStopCell(headerRow, getLabel("column.color"), COLOR_TEXT_WIDTH, false);
        return {
            headerRow: headerRow,
            locationLabel: addStopCell(headerRow, getLabel("column.location"), VALUE_COLUMN_WIDTH, false),
            opacityLabel: addStopCell(headerRow, getLabel("column.opacity"), VALUE_COLUMN_WIDTH, false)
        };
    }

    /**
     * 位置・不透明度の見出しを、1行目の入力欄の左端に揃える（レイアウトが決まったあと、show のときに呼ぶ）。
     * 入力欄の前に∧∨があり、欄の幅は文字数で決まるので、作るときには位置を決められない
     * @param {Object} headerControls - addStopHeaderRow() の戻り値
     * @param {Object} firstRow - 同じパネルの1行目（addStopRow() の戻り値）
     * @returns {void}
     */
    function alignHeaderLabels(headerControls, firstRow) {
        alignLabelToInput(headerControls.locationLabel, firstRow.locationInput, headerControls.headerRow);
        alignLabelToInput(headerControls.opacityLabel, firstRow.opacityInput, headerControls.headerRow);
    }

    /**
     * 見出しを入力欄の左端の真上（欄の中の文字の書き出しに合わせて HEADER_LABEL_INSET だけ右）に動かし、幅を文字に合わせる
     * @param {StaticText} headerLabel - 見出し
     * @param {EditText} valueInput - 入力欄
     * @param {Group} headerRow - 見出しの行（入力欄の行と同じパネルにある）
     * @returns {void}
     */
    function alignLabelToInput(headerLabel, valueInput, headerRow) {
        if (!headerLabel.size || !valueInput.size) return;
        var stopPanel = headerRow.parent;
        var inputLeft = getOffsetInAncestor(valueInput, stopPanel) - getOffsetInAncestor(headerRow, stopPanel);
        var textWidth = Math.ceil(headerLabel.graphics.measureString(headerLabel.text)[0]) + 2;
        headerLabel.size = [textWidth, headerLabel.size.height];
        headerLabel.location = [inputLeft + HEADER_LABEL_INSET, headerLabel.location[1]];
    }

    /**
     * コントロールの左端が、先祖のコンテナの中で何 px の位置にあるかを返す
     * @param {Object} control - コントロール
     * @param {Object} ancestor - 先祖のコンテナ
     * @returns {number} 横の位置（px）
     */
    function getOffsetInAncestor(control, ancestor) {
        var offsetX = 0;
        for (var currentNode = control; currentNode && currentNode !== ancestor; currentNode = currentNode.parent) offsetX += currentNode.location[0];
        return offsetX;
    }

    /**
     * ストップ1つの行（見本・カラー・位置・不透明度）を足す。見本のクリックでカラーを選び直す
     * @param {Panel} parent - 追加先のパネル
     * @param {Object} stopEntry - readGradientEntry() のストップ
     * @param {Function} onStopEdited - 値が変わったときに呼ぶ関数
     * @returns {Object} { rowGroup, stopEntry, colorChip, colorText, locationInput, opacityInput }（readStopRowEdits() で読む）
     */
    function addStopRow(parent, stopEntry, onStopEdited) {
        var stopRow = addStopRowGroup(parent);
        /* 作り直したときも、変更中の値で表示する / show the edited values when rows are rebuilt */
        var currentColor = stopEntry.editedColor || stopEntry.color;
        var currentOpacity = (stopEntry.editedOpacity !== null) ? stopEntry.editedOpacity : stopEntry.opacity;
        var colorChip = addColorChip(stopRow, toScreenRgb(currentColor));
        var colorText = addStopCell(stopRow, describeColor(currentColor), COLOR_TEXT_WIDTH, false);
        colorChip.helpTip = getLabel("tooltip.colorChip");
        var stopRowControls = {
            rowGroup: stopRow,
            stopEntry: stopEntry,
            colorChip: colorChip,
            colorText: colorText,
            locationInput: addPercentField(stopRow, stopEntry.rampPoint, getCurrentRampPoint(stopEntry), onStopEdited),
            opacityInput: addPercentField(stopRow, stopEntry.opacity, currentOpacity, onStopEdited)
        };
        colorChip.addEventListener("click", function () {
            if (pickStopColor(stopRowControls)) onStopEdited();
        });
        return stopRowControls;
    }

    /**
     * 行の末尾に、上下のストップと入れ替える矢印（先頭の↑・末尾の↓は無効）と、複製・削除のボタンを足す（ストップが2つなら削除は無効）。
     * 値を増減する∧∨と見分けられるよう、枠の無いアイコンを横に並べる
     * @param {Object} panelControls - addGradientPanel() の戻り値
     * @param {number} rowIndex - 対象の行の番号
     * @returns {void}
     */
    function addRowButtons(panelControls, rowIndex) {
        var panelRows = panelControls.rows;
        var stopRowControls = panelRows[rowIndex];
        var onStopEdited = panelControls.panelCallbacks.onStopEdited;
        var rowButtonGroup = stopRowControls.rowGroup.add("group");
        setupRow(rowButtonGroup, "left", ROW_BUTTON_SPACING);
        rowButtonGroup.margins = [ROW_BUTTON_GROUP_LEFT_MARGIN, 0, 0, 0]; /* 値の欄と離して、別の操作だと分かるようにする / keep apart from the value fields */
        var moveUpButton = addIconButton(rowButtonGroup, ROW_BUTTON_SIZE, drawMoveUpIcon);
        var moveDownButton = addIconButton(rowButtonGroup, ROW_BUTTON_SIZE, drawMoveDownIcon);
        var duplicateButton = addIconButton(rowButtonGroup, ROW_BUTTON_SIZE, drawDuplicateIcon);
        var removeButton = addIconButton(rowButtonGroup, ROW_BUTTON_SIZE, drawRemoveIcon);
        moveUpButton.helpTip = getLabel("tooltip.swapUp");
        moveDownButton.helpTip = getLabel("tooltip.swapDown");
        duplicateButton.helpTip = getLabel("tooltip.duplicate");
        removeButton.helpTip = getLabel("tooltip.removeStop");
        moveUpButton.onClick = function () {
            swapStopRows(panelRows[rowIndex - 1], stopRowControls);
            onStopEdited();
        };
        moveDownButton.onClick = function () {
            swapStopRows(stopRowControls, panelRows[rowIndex + 1]);
            onStopEdited();
        };
        duplicateButton.onClick = function () {
            duplicateStopRow(panelControls, rowIndex);
        };
        removeButton.onClick = function () {
            removeStopRow(panelControls, rowIndex);
        };
        setIconButtonEnabled(moveUpButton, rowIndex > 0);
        setIconButtonEnabled(moveDownButton, rowIndex < panelRows.length - 1);
        setIconButtonEnabled(removeButton, panelRows.length > MIN_STOP_COUNT);
    }

    /**
     * 削除のアイコン（×）を描く
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawRemoveIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        var halfSize = REMOVE_ICON_SIZE / 2;
        var centerX = iconWidth / 2;
        var centerY = iconHeight / 2;
        iconGraphics.newPath();
        iconGraphics.moveTo(centerX - halfSize, centerY - halfSize);
        iconGraphics.lineTo(centerX + halfSize, centerY + halfSize);
        iconGraphics.moveTo(centerX + halfSize, centerY - halfSize);
        iconGraphics.lineTo(centerX - halfSize, centerY + halfSize);
        iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, REMOVE_ICON_LINE_WIDTH));
    }

    /**
     * 複製のアイコン（ずらして重ねた2つの四角）を描く
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawDuplicateIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        var squareSize = DUPLICATE_ICON_SQUARE_SIZE;
        var shift = DUPLICATE_ICON_SHIFT;
        var backLeft = Math.round((iconWidth - squareSize - shift) / 2) + 0.5;
        var backTop = Math.round((iconHeight - squareSize - shift) / 2) + 0.5;
        var iconPen = iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, DUPLICATE_ICON_LINE_WIDTH);
        /* 後ろの四角は、前の四角に隠れる右下を描かない / the back square skips the corner hidden by the front one */
        iconGraphics.newPath();
        iconGraphics.moveTo(backLeft, backTop + squareSize);
        iconGraphics.lineTo(backLeft, backTop);
        iconGraphics.lineTo(backLeft + squareSize, backTop);
        iconGraphics.strokePath(iconPen);
        iconGraphics.newPath();
        iconGraphics.rectPath(backLeft + shift, backTop + shift, squareSize, squareSize);
        iconGraphics.strokePath(iconPen);
    }

    /**
     * 上向きの矢印（↑）を描く
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawMoveUpIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        drawVerticalArrowIcon(iconGraphics, iconWidth, iconHeight, iconColor, -1);
    }

    /**
     * 下向きの矢印（↓）を描く
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawMoveDownIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        drawVerticalArrowIcon(iconGraphics, iconWidth, iconHeight, iconColor, 1);
    }

    /**
     * 軸と矢じりのある縦の矢印を線で描く
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @param {number} direction - 上向きなら -1、下向きなら 1
     * @returns {void}
     */
    function drawVerticalArrowIcon(iconGraphics, iconWidth, iconHeight, iconColor, direction) {
        var centerX = iconWidth / 2;
        var centerY = iconHeight / 2;
        var halfLength = SWAP_ARROW_LENGTH / 2;
        var tipY = centerY + direction * halfLength; /* 矢じりの先 / arrow tip */
        var arrowPen = iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, SWAP_ARROW_LINE_WIDTH);
        iconGraphics.newPath();
        iconGraphics.moveTo(centerX, centerY - direction * halfLength);
        iconGraphics.lineTo(centerX, tipY);
        iconGraphics.moveTo(centerX - SWAP_ARROW_HEAD_SIZE, tipY - direction * SWAP_ARROW_HEAD_SIZE);
        iconGraphics.lineTo(centerX, tipY);
        iconGraphics.lineTo(centerX + SWAP_ARROW_HEAD_SIZE, tipY - direction * SWAP_ARROW_HEAD_SIZE);
        iconGraphics.strokePath(arrowPen);
    }

    /**
     * 2つの行のカラーと不透明度を入れ替える（位置はそのまま）
     * @param {Object} upperRow - 上の行（addStopRow() の戻り値）
     * @param {Object} lowerRow - 下の行
     * @returns {void}
     */
    function swapStopRows(upperRow, lowerRow) {
        var upperColor = getRowColor(upperRow);
        setRowColor(upperRow, getRowColor(lowerRow));
        setRowColor(lowerRow, upperColor);
        var upperOpacityText = upperRow.opacityInput.text;
        upperRow.opacityInput.text = lowerRow.opacityInput.text;
        lowerRow.opacityInput.text = upperOpacityText;
        /* 不正な入力で戻す先も入れ替える / keep the revert value in step */
        upperRow.opacityInput.lastValidText = upperRow.opacityInput.text;
        lowerRow.opacityInput.lastValidText = lowerRow.opacityInput.text;
    }

    /**
     * 行の今のカラーを返す（選び直していればそのカラー）
     * @param {Object} stopRowControls - addStopRow() の戻り値
     * @returns {Color} カラー
     */
    function getRowColor(stopRowControls) {
        return stopRowControls.stopEntry.editedColor || stopRowControls.stopEntry.color;
    }

    /**
     * 行のカラーを変え、見本と文字を更新する
     * @param {Object} stopRowControls - addStopRow() の戻り値
     * @param {Color} color - 新しいカラー
     * @returns {void}
     */
    function setRowColor(stopRowControls, color) {
        stopRowControls.stopEntry.editedColor = color;
        stopRowControls.colorText.text = describeColor(color);
        stopRowControls.colorChip.rgbValues = toScreenRgb(color);
        /* group には notify() が無いので、隠して再表示して描き直す / groups have no notify(), so hide and show to redraw */
        stopRowControls.colorChip.hide();
        stopRowControls.colorChip.show();
    }

    /**
     * 0〜100% の数値欄（∧∨付き）を、決まった幅の欄に右揃えで足す
     * @param {Group} stopRow - 行
     * @param {number} baseValue - 元の値（変更の有無はこの値の文字列と比べる）
     * @param {number} currentValue - 表示する値
     * @param {Function} onStopEdited - 値が変わったとき（∧∨・↑↓・確定）に呼ぶ関数
     * @returns {EditText} 入力欄（.initialText に元の値の文字列を控える）
     */
    function addPercentField(stopRow, baseValue, currentValue, onStopEdited) {
        var valueCell = stopRow.add("group");
        valueCell.orientation = "row";
        valueCell.alignChildren = ["right", "center"];
        valueCell.margins = 0;
        valueCell.preferredSize.width = VALUE_COLUMN_WIDTH;
        var valueInput = addSteppedField(valueCell, {
            text: formatSteppedValue(currentValue, PERCENT_FIELD_OPTIONS),
            characters: PERCENT_FIELD_CHARACTERS,
            step: PERCENT_FIELD_OPTIONS.step,
            min: PERCENT_FIELD_OPTIONS.min,
            max: PERCENT_FIELD_OPTIONS.max,
            unit: PERCENT_FIELD_OPTIONS.unit,
            onStep: onStopEdited /* ∧∨と↑↓キー / stepper and arrow keys */
        });
        valueInput.initialText = formatSteppedValue(baseValue, PERCENT_FIELD_OPTIONS);
        /* 直接入力の確定 / typed values on commit */
        valueInput.addEventListener("change", onStopEdited);
        return valueInput;
    }

    /**
     * 行の数値欄から変更を読み、ストップの editedRampPoint / editedOpacity に入れる（文字が初期のままなら変えない）
     * @param {Object} stopRowControls - addStopRow() の戻り値
     * @returns {void}
     */
    function readStopRowEdits(stopRowControls) {
        stopRowControls.stopEntry.editedRampPoint = readPercentField(stopRowControls.locationInput);
        stopRowControls.stopEntry.editedOpacity = readPercentField(stopRowControls.opacityInput);
    }

    /**
     * 数値欄の値を読む（初期のままか、数値として読めなければ null）
     * @param {EditText} valueInput - addPercentField() の入力欄
     * @returns {number|null} 0〜100 の値
     */
    function readPercentField(valueInput) {
        if (valueInput.text === valueInput.initialText) return null;
        var percentValue = evaluateArithmetic(valueInput.text, PERCENT_FIELD_OPTIONS.unit);
        if (isNaN(percentValue)) return null;
        return clampSteppedValue(percentValue, PERCENT_FIELD_OPTIONS);
    }

    /**
     * 標準のカラーピッカーでストップのカラーを選び直し、見本と文字を更新する
     * @param {Object} stopRowControls - addStopRow() の戻り値
     * @returns {boolean} カラーを変えたら true
     */
    function pickStopColor(stopRowControls) {
        var currentColor = getRowColor(stopRowControls);
        var pickedColor;
        /* スウォッチのカラーなど、ピッカーが受け付けない色では例外になることがある / the picker may refuse some colors, such as swatch colors */
        try {
            pickedColor = app.showColorPicker(currentColor);
        } catch (e) {
            return false;
        }
        /* キャンセルでは渡した色がそのまま返る / Cancel returns the color that was passed in */
        if (!pickedColor || describeColor(pickedColor) === describeColor(currentColor)) return false;
        setRowColor(stopRowControls, pickedColor);
        return true;
    }

    /**
     * カラーの見本を足す（変換できない色は斜線）。色は .rgbValues を書き換えて描き直す
     * @param {Group} parent - 追加先の行
     * @param {number[]|null} rgbValues - [R, G, B]（0〜255）
     * @returns {Group} 見本
     */
    function addColorChip(parent, rgbValues) {
        var colorChip = parent.add("group");
        colorChip.preferredSize = [COLOR_CHIP_SIZE, COLOR_CHIP_SIZE];
        colorChip.minimumSize = [COLOR_CHIP_SIZE, COLOR_CHIP_SIZE];
        colorChip.maximumSize = [COLOR_CHIP_SIZE, COLOR_CHIP_SIZE];
        colorChip.rgbValues = rgbValues;
        var frameColor = isDarkUI() ? [1, 1, 1, 0.35] : [0, 0, 0, 0.35];
        colorChip.onDraw = function () {
            var chipGraphics = colorChip.graphics;
            var chipWidth = colorChip.size.width;
            var chipHeight = colorChip.size.height;
            var chipRgb = colorChip.rgbValues;
            var framePen = chipGraphics.newPen(chipGraphics.PenType.SOLID_COLOR, frameColor, 1);
            if (chipRgb) {
                chipGraphics.newPath();
                chipGraphics.rectPath(0, 0, chipWidth, chipHeight);
                chipGraphics.fillPath(chipGraphics.newBrush(chipGraphics.BrushType.SOLID_COLOR,
                    [chipRgb[0] / 255, chipRgb[1] / 255, chipRgb[2] / 255, 1]));
            } else {
                chipGraphics.newPath();
                chipGraphics.moveTo(chipWidth, 0);
                chipGraphics.lineTo(0, chipHeight);
                chipGraphics.strokePath(framePen);
            }
            chipGraphics.newPath();
            chipGraphics.rectPath(0.5, 0.5, chipWidth - 1, chipHeight - 1);
            chipGraphics.strokePath(framePen);
        };
        return colorChip;
    }

    // -----------------------------------------
    // 境界のダイアログ / Boundary dialog
    // -----------------------------------------

    /**
     * セパレートの境界の位置を入力するダイアログを表示する。onPreview を渡すと、入力に合わせて反映する（プレビューは常にオン）
     * @param {Object} dialogSetup - 表示の設定
     * @param {string} dialogSetup.modeText - 上に出す説明
     * @param {number[]} dialogSetup.initialBoundaries - 境界の初期値（%）
     * @param {boolean} dialogSetup.hasFooter - 左下にスクリプト名・バージョンを出すなら true（一覧のダイアログから開く小さなダイアログは false）
     * @param {function(number[]): void} [dialogSetup.onPreview] - 使える境界が入力されるたびに呼ぶ（取り消しは呼び出し側で戻す）
     * @param {{label: string, helpTip: string, onClick: Function}|null} [dialogSetup.extraButton] - ボタン行の左に置くボタン
     * @returns {number[]|null} 境界の位置（キャンセルなら null）
     */
    function showBoundaryDialog(dialogSetup) {
        var boundaryDialog = new Window("dialog", getLabel("dialog.separateTitle"));
        setupWindow(boundaryDialog);

        boundaryDialog.add("statictext", undefined, dialogSetup.modeText);

        var boundaryPanel = boundaryDialog.add("panel", undefined, getLabel("panel.boundaries"));
        setupPanel(boundaryPanel, 6);

        var boundaryInputs = [];
        for (var i = 0; i < dialogSetup.initialBoundaries.length; i++) {
            var boundaryInput = addSteppedField(boundaryPanel, {
                label: labelText("fieldLabel.boundary", [i + 1]),
                labelWidth: BOUNDARY_LABEL_WIDTH,
                text: formatStepperNumber(dialogSetup.initialBoundaries[i]) + BOUNDARY_FIELD_OPTIONS.unit,
                characters: BOUNDARY_FIELD_CHARACTERS,
                step: BOUNDARY_FIELD_OPTIONS.step,
                min: BOUNDARY_FIELD_OPTIONS.min,
                max: BOUNDARY_FIELD_OPTIONS.max,
                unit: BOUNDARY_FIELD_OPTIONS.unit,
                onStep: refreshPreview
            });
            boundaryInput.helpTip = getLabel("tooltip.boundary");
            appendOnChange(boundaryInput, refreshPreview);
            boundaryInputs.push(boundaryInput);
        }

        var buttonRow = addButtonRow(boundaryDialog);
        if (dialogSetup.extraButton) {
            var extraButton = buttonRow.leftGroup.add("button", undefined, dialogSetup.extraButton.label);
            extraButton.helpTip = dialogSetup.extraButton.helpTip;
            extraButton.onClick = dialogSetup.extraButton.onClick;
        }
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var okButton = buttonRow.rightGroup.add("button", undefined, "OK", { name: "ok" });

        /**
         * 入力欄の値を読む
         * @returns {number[]} 境界の位置（%。読めない欄は NaN）
         */
        function readBoundaries() {
            var boundaries = [];
            for (var j = 0; j < boundaryInputs.length; j++) boundaries.push(parseFloat(boundaryInputs[j].text));
            return boundaries;
        }

        /**
         * 入力した境界をプレビューに反映する（使えない値のあいだは前の表示のまま）
         * @returns {void}
         */
        function refreshPreview() {
            if (!dialogSetup.onPreview) return;
            var boundaries = readBoundaries();
            if (areBoundariesValid(boundaries)) dialogSetup.onPreview(boundaries);
        }

        okButton.onClick = function () {
            if (!areBoundariesValid(readBoundaries())) {
                alert(getLabel("alert.invalidBoundaries"));
                return;
            }
            boundaryDialog.close(1);
        };

        refreshPreview();
        alignRightOnlyButtonRow(buttonRow);
        if (dialogSetup.hasFooter) addVersionFooter(boundaryDialog);
        prepareDialogWindow(boundaryDialog, SCRIPT_NAME + "Boundary");
        if (boundaryDialog.show() !== 1) return null;
        return readBoundaries();
    }

    /**
     * 入力欄の確定時の処理に、あとから処理を足す（StepperButtons の onChange で値をそろえたあとに呼ぶ）
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {Function} afterChange - 足す処理
     * @returns {void}
     */
    function appendOnChange(numberInput, afterChange) {
        var baseOnChange = numberInput.onChange;
        numberInput.onChange = function () {
            if (baseOnChange) baseOnChange();
            afterChange();
        };
    }

})();
