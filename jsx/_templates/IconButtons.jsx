#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

保存（トレイに下向きの矢印）・削除（ゴミ箱）・自動計算（魔法の杖ときらめき）のアイコンのボタンを、onDraw で描く再利用テンプレートです。
プリセットの保存・削除や、値を自動で求める［自動］など、文字のボタンの代わりに使います。無効の間は押せず、薄く描きます。

### Overview

A reusable template for save (tray with a down arrow), delete (trash can), and auto (magic wand with sparkles) icon buttons drawn in onDraw.
Use them in place of text buttons, e.g. to save or delete presets or to calculate a value automatically. While disabled they cannot be pressed and are drawn faint.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IconButtons";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-20";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // 【移植手順 / How to port】
    // 1. 「（再利用パーツ）」の行から「ここまで」の行までをまるごと、コピー先の IIFE 内に貼る。
    //    識別子は ICON_BUTTON_* / *IconButton* / drawSaveIcon / drawTrashIcon / drawAutoIcon
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の部品も貼っておく）
    // 2. アイコンを addIconButton(親, [幅, 高さ], 描く関数) で作り、onClick と helpTip はコピー先で付ける
    //      var presetSaveIcon = addIconButton(presetRow, ICON_BUTTON_SIZE, drawSaveIcon);
    //      presetSaveIcon.helpTip = getLabel("tooltip.presetSave");
    //      presetSaveIcon.onClick = function () { … };
    // 3. 有効／無効は setIconButtonEnabled(iconButton, isEnabled)（自作描画は enabled を変えるだけでは描き直されない）
    // 4. 描く関数は (iconGraphics, iconWidth, iconHeight, iconColor) の形。ほかの絵を足すときも同じ形で書く

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

    /**
     * デモ用：UI がダークテーマかどうか（本番では UITheme の部品を貼る）
     * @returns {boolean} ダークなら true
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

    // =========================================
    // メイン処理（デモ。コピー時は不要） / Main (demo; omit when copying)
    // =========================================

    /**
     * ドロップダウンの右に保存・削除のアイコンを並べたデモダイアログを表示する。
     * 「---」を選んでいる間は削除をディムにする
     * @returns {void}
     */
    function showDemoDialog() {
        var isJapanese = ($.locale.indexOf("ja") === 0);
        var demoDialog = new Window("dialog", SCRIPT_NAME + " " + SCRIPT_VERSION);
        demoDialog.orientation = "column";
        demoDialog.alignChildren = ["fill", "top"];
        demoDialog.margins = 16;

        var presetRow = demoDialog.add("group");
        presetRow.orientation = "row";
        presetRow.alignChildren = ["left", "center"];
        var presetDropdown = presetRow.add("dropdownlist", undefined, ["---", "A", "B"]);
        presetDropdown.selection = 0;
        var saveIcon = addIconButton(presetRow, ICON_BUTTON_SIZE, drawSaveIcon);
        saveIcon.helpTip = isJapanese ? "保存" : "Save";
        saveIcon.onClick = function () { alert(isJapanese ? "保存" : "Save"); };
        var deleteIcon = addIconButton(presetRow, ICON_BUTTON_SIZE, drawTrashIcon);
        deleteIcon.helpTip = isJapanese ? "削除" : "Delete";
        deleteIcon.onClick = function () { alert((isJapanese ? "削除 : " : "Delete: ") + presetDropdown.selection.text); };
        setIconButtonEnabled(deleteIcon, false);
        presetDropdown.onChange = function () { setIconButtonEnabled(deleteIcon, presetDropdown.selection.index > 0); };

        /* 数値欄の右に自動計算のアイコン / an Auto icon next to a number field */
        var countRow = demoDialog.add("group");
        countRow.orientation = "row";
        countRow.alignChildren = ["left", "center"];
        var countInput = countRow.add("edittext", undefined, "3");
        countInput.characters = 4;
        var autoIcon = addIconButton(countRow, ICON_BUTTON_SIZE, drawAutoIcon);
        autoIcon.helpTip = isJapanese ? "自動計算" : "Calculate automatically";
        autoIcon.onClick = function () { countInput.text = String(Math.floor(Math.random() * 9) + 1); };

        demoDialog.add("button", undefined, "OK", { name: "ok" });
        demoDialog.show();
    }

    showDemoDialog();

})();
