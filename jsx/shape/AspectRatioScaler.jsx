#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、指定した縦横比に合わせてサイズ変更します。
固定する辺の長さと基準点も指定でき、選択がないときはその比率の長方形を作ります。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AspectRatioScaler.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n4a212e6eacf1

### Overview

Resizes the selected objects to a chosen aspect ratio.
You can also set the fixed side's length and the reference point; with nothing selected, it draws a rectangle of that ratio.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AspectRatioScaler.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AspectRatioScaler";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AspectRatioScaler.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AspectRatioScaler.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n4a212e6eacf1"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* プリセットの比率（横 ÷ 縦）/ Preset ratios (width / height) */
    var RATIO_16_9   = 16 / 9;
    var RATIO_SQUARE = 1;
    var RATIO_A4     = 210 / 297;

    /* カスタム比率の初期値 / Initial custom ratio */
    var DEFAULT_CUSTOM_RATIO_WIDTH  = "3";
    var DEFAULT_CUSTOM_RATIO_HEIGHT = "2";

    /* 選択なしで作る長方形の、固定する辺の欄が空のときの長さ（pt）/ Fixed-side length of the rectangle drawn with nothing selected, when the field is empty */
    var FALLBACK_BASE_SIZE_PT = 200;

    /* 選択なしのとき幅の欄に入れる初期値（単位コード → 値）/ Width field default when nothing is selected (unit code -> value) */
    var DEFAULT_SIZE_TEXT_BY_UNIT = { 1: "100", 6: "1000" };

    /* 基準点の初期値（0..8 を行優先、0=左上・4=中央・8=右下）/ Initial reference point (row-major 0..8; 0=top-left, 4=center, 8=bottom-right) */
    var DEFAULT_ANCHOR_INDEX = 4;

    /* ［ピクセルグリッドに最適化］の初期状態 / Initial state of "Make Pixel Perfect" */
    var DEFAULT_ALIGN_TO_PIXEL_GRID = true;

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS       = [15, 20, 15, 10];  /* パネル余白 [左,上,右,下] */
    var FIELD_CHARACTERS    = 5;                 /* 数値欄の幅（文字数）/ Numeric field width */
    var DIALOG_OFFSET_X     = 300;               /* ダイアログを右へずらす量 / Horizontal dialog offset */
    var DIALOG_OPACITY      = 0.97;              /* ダイアログの不透明度 / Dialog opacity */
    var ANCHOR_WIDGET_SIZE  = 66;                /* 9軸ウィジェット全体の大きさ / Overall size of the 9-axis widget */
    var ANCHOR_CELL_SIZE    = 9;                 /* 9軸の□1個のサイズ / Size of one anchor square */
    var ANCHOR_CELL_GAP     = 7.5;               /* 9軸の□どうしの間隔 / Gap between anchor squares */
    var CUSTOM_RATIO_INDENT = 14;                /* カスタム比の欄の左インデント（約1文字）/ Left indent of the custom ratio fields (about one character) */
    var ORIENT_BUTTON_SIZE  = 36;                /* 向きアイコンのボタンの大きさ / Size of an orientation icon button */
    var ORIENT_FRAME_LONG   = 30;                /* 向きアイコンの枠の長辺 / Long side of the orientation icon frame */
    var ORIENT_FRAME_SHORT  = 23;                /* 向きアイコンの枠の短辺 / Short side of the orientation icon frame */

    /**
     * 見出し付きパネルを縦並びで追加する
     * @param {Group} parent - 追加先
     * @param {string} title - パネルの見出し
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, title) {
        var titledPanel = parent.add("panel", undefined, title);
        titledPanel.orientation = "column";
        titledPanel.alignChildren = "left";
        titledPanel.margins = PANEL_MARGINS;
        titledPanel.alignment = ["fill", "top"];
        return titledPanel;
    }

    /**
     * 子を縦に並べるグループを追加する
     * @param {Object} parent - 追加先のパネルまたはグループ
     * @returns {Group} 追加したグループ
     */
    function addColumnGroup(parent) {
        var columnGroup = parent.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = "left";
        return columnGroup;
    }

    /**
     * 子を横に並べ、天地中央でそろえるグループを追加する
     * @param {Object} parent - 追加先のパネルまたはグループ
     * @returns {Group} 追加したグループ
     */
    function addRowGroup(parent) {
        var rowGroup = parent.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignment = ["left", "top"];
        rowGroup.alignChildren = ["left", "center"];
        return rowGroup;
    }

    /**
     * ダイアログを表示時に横へずらす
     * @param {Window} dialog - 対象のダイアログ
     * @param {number} offsetX - 横方向のずらし量
     * @param {number} offsetY - 縦方向のずらし量
     * @returns {void}
     */
    function shiftDialogPosition(dialog, offsetX, offsetY) {
        dialog.onShow = function () {
            dialog.location = [dialog.location[0] + offsetX, dialog.location[1] + offsetY];
        };
    }

    /**
     * 上下キーで数値を増減する（shift: ±10、option: ±0.1）
     * @param {EditText} editText - 対象の入力欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName != "Up" && event.keyName != "Down") return;
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Optionキー押下時は0.1単位で増減
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                // 小数第1位までに丸め
                value = Math.round(value * 10) / 10;
            } else {
                // 整数に丸め
                value = Math.round(value);
            }

            editText.text = value;
            if (typeof editText.onChanging === "function") editText.onChanging();
        });
    }

    // =========================================
    // 自作描画のウィジェット共通 / Custom-drawn widgets (shared)
    // =========================================

    /* UI が明るいテーマか / Whether the UI uses a light theme */
    var IS_LIGHT_UI = (app.preferences.getRealPreference("uiBrightness") > 0.5);

    /**
     * 矩形を塗る（多角形は fillPath で塗れないので rectPath を使う）
     * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
     * @param {number[]} rect - [x, y, 幅, 高さ]
     * @param {number[]} color - RGBA
     * @returns {void}
     */
    function fillRect(graphics, rect, color) {
        graphics.newPath();
        graphics.rectPath(rect[0], rect[1], rect[2], rect[3]);
        graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
    }

    /**
     * 矩形の輪郭を描く
     * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
     * @param {number[]} rect - [x, y, 幅, 高さ]
     * @param {number[]} color - RGBA
     * @param {number} lineWidth - 線幅
     * @returns {void}
     */
    function strokeRect(graphics, rect, color, lineWidth) {
        graphics.newPath();
        graphics.rectPath(rect[0], rect[1], rect[2], rect[3]);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, lineWidth));
    }

    /**
     * 楕円を塗る
     * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
     * @param {number[]} rect - 外接矩形 [x, y, 幅, 高さ]
     * @param {number[]} color - RGBA
     * @returns {void}
     */
    function fillEllipse(graphics, rect, color) {
        graphics.newPath();
        graphics.ellipsePath(rect[0], rect[1], rect[2], rect[3]);
        graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
    }

    /**
     * コントロールの地をパネルと同じ色で塗って透過に見せる
     * @param {Object} control - 対象のコントロール
     * @returns {void}
     */
    function paintControlBackground(control) {
        var graphics = control.graphics;
        /* backgroundColor が無い環境では例外 / Throws where backgroundColor is missing */
        try {
            graphics.newPath();
            graphics.rectPath(0, 0, control.size[0], control.size[1]);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) {}
    }

    /**
     * コントロールを再描画する
     * @param {Object} control - 対象のコントロール
     * @returns {void}
     */
    function redrawControl(control) {
        /* notify("onDraw") は環境により例外になる / Throws on some versions */
        try { control.notify("onDraw"); } catch (e) {}
    }

    // =========================================
    // 9軸ウィジェット / Anchor widget
    // =========================================

    /* セルの枠線とケイ線：薄いグレー / Cell borders and connecting rules: light gray */
    var ANCHOR_LINE_COLOR = [0.6, 0.6, 0.6, 1];
    /* 選択セルの塗り：ライトUIは濃いグレー、ダークUIは明るいグレー / Selected-cell fill: dark gray in light UI, bright gray in dark UI */
    var ANCHOR_SELECTED_FILL = IS_LIGHT_UI ? [0.4, 0.4, 0.4, 1] : [0.8, 0.8, 0.8, 1];
    /* 中央(4)を除く外周の□どうしをつなぐケイ線の組み合わせ / Pairs of outer squares (center 4 excluded) joined by rules */
    var ANCHOR_CONNECTIONS = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];

    /**
     * 基準点セルの□を1つ描画する（選択時だけ塗る）
     * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
     * @param {number} x - 左端
     * @param {number} y - 上端
     * @param {boolean} selected - 選択中なら true
     * @returns {void}
     */
    function drawAnchorCell(graphics, x, y, selected) {
        var cellRect = [x, y, ANCHOR_CELL_SIZE, ANCHOR_CELL_SIZE];
        /* 枠を上に描くので塗りを先に行う / Fill first so the border draws on top */
        if (selected) fillRect(graphics, cellRect, ANCHOR_SELECTED_FILL);
        strokeRect(graphics, cellRect, ANCHOR_LINE_COLOR, 1);
    }

    /**
     * 9軸ウィジェットを描画する（外周の□をケイ線でつなぐ・中央は独立）
     * @param {Button} widget - 対象のウィジェット
     * @returns {void}
     */
    function drawAnchorWidget(widget) {
        var graphics = widget.graphics;
        var width = widget.size[0];
        var height = widget.size[1];

        paintControlBackground(widget);

        var cellStep = ANCHOR_CELL_SIZE + ANCHOR_CELL_GAP;
        var gridSize = ANCHOR_CELL_SIZE * 3 + ANCHOR_CELL_GAP * 2;
        var originX = Math.round((width - gridSize) / 2);
        var originY = Math.round((height - gridSize) / 2);

        /* 9セルの左上座標を先に求める / Precompute the top-left corner of all nine cells */
        var cellPositions = [];
        for (var i = 0; i < 9; i++) {
            cellPositions.push([originX + (i % 3) * cellStep, originY + Math.floor(i / 3) * cellStep]);
        }

        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, ANCHOR_LINE_COLOR, 1);
        for (var j = 0; j < ANCHOR_CONNECTIONS.length; j++) {
            var cellA = cellPositions[ANCHOR_CONNECTIONS[j][0]];
            var cellB = cellPositions[ANCHOR_CONNECTIONS[j][1]];
            graphics.newPath();
            if (ANCHOR_CONNECTIONS[j][1] - ANCHOR_CONNECTIONS[j][0] === 1) {
                /* 横方向：右隣の□へ / Horizontal: to the square on the right */
                graphics.moveTo(cellA[0] + ANCHOR_CELL_SIZE, cellA[1] + ANCHOR_CELL_SIZE / 2);
                graphics.lineTo(cellB[0], cellB[1] + ANCHOR_CELL_SIZE / 2);
            } else {
                /* 縦方向：下の□へ / Vertical: to the square below */
                graphics.moveTo(cellA[0] + ANCHOR_CELL_SIZE / 2, cellA[1] + ANCHOR_CELL_SIZE);
                graphics.lineTo(cellB[0] + ANCHOR_CELL_SIZE / 2, cellB[1]);
            }
            graphics.strokePath(linePen);
        }

        for (var k = 0; k < cellPositions.length; k++) {
            drawAnchorCell(graphics, cellPositions[k][0], cellPositions[k][1], k === widget.anchorIndex);
        }
    }

    /**
     * 9軸（3×3）の基準点ウィジェットを追加する。選び直すと widget.onAnchorChange() を呼ぶ
     * @param {Object} parent - 追加先
     * @param {Object} tipSet - ツールチップ
     * @returns {Button} 追加したウィジェット（anchorIndex に 0..8）
     */
    function addAnchorWidget(parent, tipSet) {
        var widget = parent.add("button", undefined, "");
        widget.helpTip = getLabel(tipSet);
        widget.minimumSize = widget.preferredSize = widget.maximumSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        widget.anchorIndex = DEFAULT_ANCHOR_INDEX;
        widget.onDraw = function () { drawAnchorWidget(this); };
        /* クリックしたセルを基準点にする（座標はコントロール基準）/ Set the anchor from the clicked cell (control-relative coords) */
        widget.addEventListener("mousedown", function (event) {
            var col = Math.min(2, Math.max(0, Math.floor(event.clientX / (widget.size[0] / 3))));
            var row = Math.min(2, Math.max(0, Math.floor(event.clientY / (widget.size[1] / 3))));
            widget.anchorIndex = row * 3 + col;
            redrawControl(widget);
            if (typeof widget.onAnchorChange === "function") widget.onAnchorChange();
        });
        return widget;
    }

    /**
     * 境界ボックス上の基準点の座標を返す
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @param {number} anchorIndex - 0..8（行優先）
     * @returns {number[]} [x, y]
     */
    function getAnchorPoint(bounds, anchorIndex) {
        var col = anchorIndex % 3;
        var row = Math.floor(anchorIndex / 3);
        return [
            bounds[0] + (bounds[2] - bounds[0]) * col / 2,
            bounds[1] + (bounds[3] - bounds[1]) * row / 2
        ];
    }

    // =========================================
    // 向きアイコン / Orientation icons
    // =========================================

    /* 未選択の枠と人物：グレー、選択中：青 / Unselected frame and figure: gray; selected: blue */
    var ORIENT_IDLE_COLOR = IS_LIGHT_UI ? [0.45, 0.45, 0.45, 1] : [0.7, 0.7, 0.7, 1];
    var ORIENT_SELECTED_COLOR = [0.22, 0.47, 0.9, 1];

    /**
     * 枠の中に胸像の人物を描く（頭は正円、肩は楕円、胴は枠の下端まで）
     * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
     * @param {number[]} frameRect - 枠 [x, y, 幅, 高さ]
     * @param {number[]} color - RGBA
     * @returns {void}
     */
    function drawPortraitFigure(graphics, frameRect, color) {
        var centerX = frameRect[0] + frameRect[2] / 2;
        var bottom = frameRect[1] + frameRect[3] - 3;
        var headSize = Math.round(frameRect[3] * 0.3);
        var headTop = frameRect[1] + Math.round(frameRect[3] * 0.18);
        var bodyWidth = Math.round(headSize * 1.9);
        var shoulderTop = headTop + headSize + 1;
        var shoulderHeight = Math.round(headSize * 0.9);

        fillEllipse(graphics, [centerX - headSize / 2, headTop, headSize, headSize], color);
        fillEllipse(graphics, [centerX - bodyWidth / 2, shoulderTop, bodyWidth, shoulderHeight], color);
        var torsoTop = shoulderTop + shoulderHeight / 2;
        fillRect(graphics, [centerX - bodyWidth / 2, torsoTop, bodyWidth, bottom - torsoTop], color);
    }

    /**
     * 向きアイコンを描画する（横長または縦長の枠＋人物）
     * @param {Button} button - 対象のボタン
     * @returns {void}
     */
    function drawOrientationButton(button) {
        var graphics = button.graphics;
        paintControlBackground(button);

        var frameWidth = button.isPortrait ? ORIENT_FRAME_SHORT : ORIENT_FRAME_LONG;
        var frameHeight = button.isPortrait ? ORIENT_FRAME_LONG : ORIENT_FRAME_SHORT;
        var frameRect = [
            Math.round((button.size[0] - frameWidth) / 2),
            Math.round((button.size[1] - frameHeight) / 2),
            frameWidth,
            frameHeight
        ];
        var color = button.isSelected ? ORIENT_SELECTED_COLOR : ORIENT_IDLE_COLOR;
        strokeRect(graphics, frameRect, color, 2);
        drawPortraitFigure(graphics, frameRect, color);
    }

    /**
     * 縦／横の向きアイコンを2つ並べて追加する（どちらか一方だけが isSelected=true）
     * 選び直すと group.onOrientationChange() を呼ぶ
     * @param {Object} parent - 追加先
     * @param {Object} portraitTip - 縦のツールチップ
     * @param {Object} landscapeTip - 横のツールチップ
     * @returns {Group} 追加したグループ（portraitButton / landscapeButton を持つ）
     */
    function addOrientationButtons(parent, portraitTip, landscapeTip) {
        var orientationGroup = addRowGroup(parent);

        /**
         * 向きアイコンのボタンを1つ追加する
         * @param {boolean} isPortrait - 縦なら true
         * @param {Object} tipSet - ツールチップ
         * @returns {Button} 追加したボタン
         */
        function addButton(isPortrait, tipSet) {
            var button = orientationGroup.add("button", undefined, "");
            button.helpTip = getLabel(tipSet);
            button.minimumSize = button.preferredSize = button.maximumSize = [ORIENT_BUTTON_SIZE, ORIENT_BUTTON_SIZE];
            button.isPortrait = isPortrait;
            button.isSelected = false;
            button.onDraw = function () { drawOrientationButton(this); };
            button.onClick = function () {
                orientationGroup.landscapeButton.isSelected = !isPortrait;
                orientationGroup.portraitButton.isSelected = isPortrait;
                redrawControl(orientationGroup.landscapeButton);
                redrawControl(orientationGroup.portraitButton);
                if (typeof orientationGroup.onOrientationChange === "function") orientationGroup.onOrientationChange();
            };
            return button;
        }

        /* 縦を左、横を右に並べる / Portrait on the left, landscape on the right */
        orientationGroup.portraitButton = addButton(true, portraitTip);
        orientationGroup.landscapeButton = addButton(false, landscapeTip);
        orientationGroup.landscapeButton.isSelected = true;
        return orientationGroup;
    }

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

    /* 定規の単位（実行中は変わらないので一度だけ読む）/ Ruler unit, read once */
    var RULER_UNIT = getUnitInfo();

    /**
     * 定規の単位に合わせて丸める（px は整数、mm は 0.1mm 刻み、その他は 0.01pt 刻み）
     * @param {number} valuePt - 値（pt）
     * @returns {number} 丸めた値（pt）
     */
    function roundForUnit(valuePt) {
        var unitCode = RULER_UNIT.code;
        if (unitCode === 6) return Math.round(valuePt); /* 1px = 1pt */
        if (unitCode === 1) {
            var stepPt = UNITS[1].pointsPerUnit * 0.1;
            return Math.round(valuePt / stepPt) * stepPt;
        }
        return Math.round(valuePt * 100) / 100;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UIの言語を返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "縦横比を指定してサイズ変更", en: "Resize to Aspect Ratio" }
        },
        panel: {
            aspectRatio: { ja: "縦横比", en: "Aspect Ratio" },
            orientationAnchor: { ja: "向きと基準点", en: "Orientation & Reference Point" },
            size: { ja: "サイズ", en: "Size" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            ratio16x9: { ja: "16:9", en: "16:9" },
            ratioSquare: { ja: "1:1（スクエア）", en: "1:1 (Square)" },
            ratioA4: { ja: "A4（1:1.414）", en: "A4 (1:1.414)" },
            ratioCustom: { ja: "カスタム", en: "Custom" },
            basisHorizontal: { ja: "幅", en: "Width" },
            basisVertical: { ja: "高さ", en: "Height" }
        },
        fieldLabel: {
            basis: { ja: "固定する辺", en: "Fixed Side" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" }
        },
        checkbox: {
            alignToPixelGrid: { ja: "ピクセルグリッドに最適化", en: "Make Pixel Perfect" },
            addArtboard: { ja: "アートボードを追加", en: "Add Artboard" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." }
        },
        tooltip: {
            anchor: {
                ja: "大きさを変えても動かない点です。クリックで選びます。",
                en: "The point that stays put when the size changes. Click to choose."
            },
            ratioPreset: {
                ja: "よく使う比率です。選ぶとカスタム欄は使いません。",
                en: "Common ratios. Selecting one disables the custom fields."
            },
            ratioCustom: { ja: "下の欄に任意の比率（横:縦）を入力します。", en: "Enter any ratio (width:height) in the fields below." },
            customWidth: { ja: "カスタム比の横の値です。", en: "The width part of the custom ratio." },
            customHeight: { ja: "カスタム比の縦の値です。", en: "The height part of the custom ratio." },
            landscape: {
                ja: "横（ランドスケープ）：長い辺を横にします（1:1 では変わりません）。",
                en: "Landscape: puts the longer side horizontally (no effect on 1:1)."
            },
            portrait: {
                ja: "縦（ポートレート）：長い辺を縦にします（1:1 では変わりません）。",
                en: "Portrait: puts the longer side vertically (no effect on 1:1)."
            },
            basisHorizontal: {
                ja: "幅を保ったまま高さを比率に合わせます。",
                en: "Keeps the width and fits the height to the ratio."
            },
            basisVertical: {
                ja: "高さを保ったまま幅を比率に合わせます。",
                en: "Keeps the height and fits the width to the ratio."
            },
            sizeValue: {
                ja: "固定する辺の長さです。空欄なら各オブジェクトの今の長さを使います。",
                en: "Length of the fixed side. Leave blank to keep each object's current length."
            },
            computedValue: {
                ja: "比率から求めた長さです。入力するには［固定する辺］をこちらに切り替えます。",
                en: "Length from the ratio. Switch the fixed side to this one to edit it."
            },
            alignToPixelGrid: {
                ja: "結果をピクセルグリッドに合わせます（［ピクセルを最適化］を実行）。",
                en: "Runs Make Pixel Perfect to align the result to the pixel grid."
            },
            addArtboard: {
                ja: "結果と同じ範囲にアートボードを追加します。オブジェクトは残ります。",
                en: "Adds an artboard matching the result. The objects stay in place."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    /**
     * 現在の言語のラベルを返す
     * @param {Object} labelSet - { ja: "...", en: "..." }
     * @returns {string} ラベル
     */
    function getLabel(labelSet) {
        return labelSet[uiLang] || labelSet.en;
    }

    /**
     * コロン付きのラベルを返す（日本語は全角、英語は半角）
     * @param {Object} labelSet - { ja: "...", en: "..." }
     * @returns {string} コロン付きのラベル
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ラジオボタンを追加する
     * @param {Group} parent - 追加先
     * @param {Object} labelSet - 表示名
     * @param {Object} tipSet - ツールチップ
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadio(parent, labelSet, tipSet) {
        var radioButton = parent.add("radiobutton", undefined, getLabel(labelSet));
        radioButton.helpTip = getLabel(tipSet);
        return radioButton;
    }

    /**
     * チェックボックスを追加する
     * @param {Object} parent - 追加先
     * @param {Object} labelSet - 表示名
     * @param {Object} tipSet - ツールチップ
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parent, labelSet, tipSet) {
        var checkbox = parent.add("checkbox", undefined, getLabel(labelSet));
        checkbox.helpTip = getLabel(tipSet);
        return checkbox;
    }

    /**
     * 数値入力欄を追加する
     * @param {Group} parent - 追加先
     * @param {string} initialText - 初期値
     * @param {Object} tipSet - ツールチップ
     * @returns {EditText} 追加した入力欄
     */
    function addNumberField(parent, initialText, tipSet) {
        var numberField = parent.add("edittext", undefined, initialText);
        numberField.helpTip = getLabel(tipSet);
        numberField.characters = FIELD_CHARACTERS;
        changeValueByArrowKey(numberField);
        return numberField;
    }

    /**
     * 「項目名：［欄］単位」の行を追加する
     * 入力欄と計算値の表示を同じ位置に重ね、setSizeRowEditable() で出し分ける
     * （無効にした入力欄は Mac で文字が薄く読めないため）
     * @param {Object} parent - 追加先
     * @param {Object} labelSet - 項目名
     * @param {number} labelWidth - 項目名の幅（右揃えでそろえる）
     * @returns {{field: EditText, valueText: StaticText}} 入力欄と計算値の表示
     */
    function addSizeRow(parent, labelSet, labelWidth) {
        var sizeRow = addRowGroup(parent);
        var rowLabel = sizeRow.add("statictext", undefined, labelText(labelSet));
        rowLabel.preferredSize.width = labelWidth;
        rowLabel.justify = "right";

        var valueStack = sizeRow.add("group");
        valueStack.orientation = "stack";
        valueStack.alignChildren = ["fill", "center"];
        var sizeField = addNumberField(valueStack, "", LABELS.tooltip.sizeValue);
        var valueText = valueStack.add("statictext", undefined, "");
        valueText.characters = FIELD_CHARACTERS;
        valueText.helpTip = getLabel(LABELS.tooltip.computedValue);

        sizeRow.add("statictext", undefined, RULER_UNIT.label);
        return { field: sizeField, valueText: valueText };
    }

    /**
     * 行を入力欄（固定する辺）か計算値の表示（固定しない辺）に切り替える
     * @param {Object} sizeRow - addSizeRow() の戻り値
     * @param {boolean} editable - 入力欄を出すなら true
     * @returns {void}
     */
    function setSizeRowEditable(sizeRow, editable) {
        sizeRow.field.visible = editable;
        sizeRow.valueText.visible = !editable;
    }

    /**
     * 「縦横比」パネルを作る（プリセットのラジオとカスタム比の欄）
     * @param {Group} parent - 追加先のカラム
     * @param {Object} dialogControls - コントロールの参照を書き込む先
     * @returns {void}
     */
    function buildRatioPanel(parent, dialogControls) {
        var ratioPanel = addPanel(parent, getLabel(LABELS.panel.aspectRatio));
        var ratioRadioGroup = addColumnGroup(ratioPanel);
        dialogControls.ratio16x9Radio = addRadio(ratioRadioGroup, LABELS.radio.ratio16x9, LABELS.tooltip.ratioPreset);
        dialogControls.ratioSquareRadio = addRadio(ratioRadioGroup, LABELS.radio.ratioSquare, LABELS.tooltip.ratioPreset);
        dialogControls.ratioA4Radio = addRadio(ratioRadioGroup, LABELS.radio.ratioA4, LABELS.tooltip.ratioPreset);
        dialogControls.ratioCustomRadio = addRadio(ratioRadioGroup, LABELS.radio.ratioCustom, LABELS.tooltip.ratioCustom);
        dialogControls.ratio16x9Radio.value = true;

        var customRatioGroup = addRowGroup(ratioPanel);
        customRatioGroup.margins = [CUSTOM_RATIO_INDENT, 0, 0, 0];
        dialogControls.customWidthField = addNumberField(customRatioGroup, DEFAULT_CUSTOM_RATIO_WIDTH, LABELS.tooltip.customWidth);
        customRatioGroup.add("statictext", undefined, ":");
        dialogControls.customHeightField = addNumberField(customRatioGroup, DEFAULT_CUSTOM_RATIO_HEIGHT, LABELS.tooltip.customHeight);
    }

    /**
     * 「オプション」パネルを作る
     * @param {Group} parent - 追加先のカラム
     * @param {Object} dialogControls - コントロールの参照を書き込む先
     * @returns {void}
     */
    function buildOptionPanel(parent, dialogControls) {
        var optionPanel = addPanel(parent, getLabel(LABELS.panel.options));
        dialogControls.alignToPixelCheckbox = addCheckbox(optionPanel, LABELS.checkbox.alignToPixelGrid, LABELS.tooltip.alignToPixelGrid);
        dialogControls.alignToPixelCheckbox.value = DEFAULT_ALIGN_TO_PIXEL_GRID;
        dialogControls.addArtboardCheckbox = addCheckbox(optionPanel, LABELS.checkbox.addArtboard, LABELS.tooltip.addArtboard);
    }

    /**
     * 「向きと基準点」パネルを作る（左に向きアイコン、右に9軸）
     * @param {Group} parent - 追加先のカラム
     * @param {Object} dialogControls - コントロールの参照を書き込む先
     * @returns {void}
     */
    function buildOrientationAnchorPanel(parent, dialogControls) {
        var orientationAnchorPanel = addPanel(parent, getLabel(LABELS.panel.orientationAnchor));
        var orientationAnchorRow = addRowGroup(orientationAnchorPanel);
        orientationAnchorRow.alignment = ["fill", "top"];
        dialogControls.orientationGroup = addOrientationButtons(orientationAnchorRow, LABELS.tooltip.portrait, LABELS.tooltip.landscape);
        dialogControls.orientationGroup.alignment = ["left", "center"]; /* 9軸と天地中央でそろえる / Vertically center with the 9-axis widget */
        dialogControls.anchorWidget = addAnchorWidget(orientationAnchorRow, LABELS.tooltip.anchor);
        dialogControls.anchorWidget.alignment = ["right", "center"];
    }

    /**
     * 「サイズ」パネルを作る（固定する辺のラジオと、幅・高さの行）
     * @param {Group} parent - 追加先のカラム
     * @param {Object} dialogControls - コントロールの参照を書き込む先
     * @returns {void}
     */
    function buildSizePanel(parent, dialogControls) {
        var sizePanel = addPanel(parent, getLabel(LABELS.panel.size));

        var basisRow = addRowGroup(sizePanel);
        basisRow.add("statictext", undefined, labelText(LABELS.fieldLabel.basis));
        dialogControls.basisHorizontalRadio = addRadio(basisRow, LABELS.radio.basisHorizontal, LABELS.tooltip.basisHorizontal);
        dialogControls.basisVerticalRadio = addRadio(basisRow, LABELS.radio.basisVertical, LABELS.tooltip.basisVertical);
        dialogControls.basisHorizontalRadio.value = true;

        /* 幅・高さの両方を表示し、固定する辺の欄だけ編集できる / Show both; only the fixed side is editable */
        var sizeLabelWidth = Math.max(
            sizePanel.graphics.measureString(labelText(LABELS.fieldLabel.width))[0],
            sizePanel.graphics.measureString(labelText(LABELS.fieldLabel.height))[0]
        );
        dialogControls.widthRow = addSizeRow(sizePanel, LABELS.fieldLabel.width, sizeLabelWidth);
        dialogControls.heightRow = addSizeRow(sizePanel, LABELS.fieldLabel.height, sizeLabelWidth);
    }

    /**
     * キャンセル・OK のボタン行を右寄せで追加する
     * @param {Window} dialog - 追加先のダイアログ
     * @returns {void}
     */
    function addButtonRow(dialog) {
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["right", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
    }

    /**
     * ダイアログを作成する（左：縦横比・オプション / 右：向きと基準点・サイズ）
     * @returns {Object} ダイアログと各コントロールの参照
     */
    function createDialog() {
        var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        dialog.opacity = DIALOG_OPACITY;
        shiftDialogPosition(dialog, DIALOG_OFFSET_X, 0);
        dialog.alignChildren = ["fill", "top"];

        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        var leftColumn = addColumnGroup(columnsGroup);
        var rightColumn = addColumnGroup(columnsGroup);
        leftColumn.alignChildren = ["fill", "top"];
        rightColumn.alignChildren = ["fill", "top"];

        var dialogControls = { dialog: dialog };
        buildRatioPanel(leftColumn, dialogControls);
        buildOptionPanel(leftColumn, dialogControls);
        buildOrientationAnchorPanel(rightColumn, dialogControls);
        buildSizePanel(rightColumn, dialogControls);
        addButtonRow(dialog);
        return dialogControls;
    }

    /**
     * ダイアログから選択中の比率（横 ÷ 縦）を読む
     * @param {Object} dialogControls - createDialog() の戻り値
     * @returns {number} 比率。カスタムが数値でないか 0 以下なら 1
     */
    function readRatio(dialogControls) {
        if (dialogControls.ratio16x9Radio.value) return RATIO_16_9;
        if (dialogControls.ratioSquareRadio.value) return RATIO_SQUARE;
        if (dialogControls.ratioA4Radio.value) return RATIO_A4;
        var ratioWidth = parseFloat(dialogControls.customWidthField.text);
        var ratioHeight = parseFloat(dialogControls.customHeightField.text);
        if (isNaN(ratioWidth) || isNaN(ratioHeight) || ratioWidth <= 0 || ratioHeight <= 0) return 1;
        return ratioWidth / ratioHeight;
    }

    /**
     * 固定する辺の欄の値を pt で返す
     * @param {Object} dialogControls - createDialog() の戻り値
     * @returns {number|null} 値（pt）。空欄・不正値・0以下なら null
     */
    function readTargetSizePt(dialogControls) {
        var fixedField = dialogControls.basisVerticalRadio.value ? dialogControls.heightRow.field : dialogControls.widthRow.field;
        var sizeValue = parseFloat(fixedField.text);
        if (isNaN(sizeValue) || sizeValue <= 0) return null;
        return sizeValue * RULER_UNIT.pointsPerUnit;
    }

    /**
     * ダイアログの設定をまとめて読む
     * @param {Object} dialogControls - createDialog() の戻り値
     * @returns {{ratio: number, wantPortrait: boolean, fixByHeight: boolean, anchorIndex: number, targetSizePt: (number|null)}} 設定
     */
    function readScaleSettings(dialogControls) {
        return {
            ratio: readRatio(dialogControls),
            wantPortrait: dialogControls.orientationGroup.portraitButton.isSelected,
            fixByHeight: dialogControls.basisVerticalRadio.value,
            anchorIndex: dialogControls.anchorWidget.anchorIndex,
            targetSizePt: readTargetSizePt(dialogControls)
        };
    }

    // =========================================
    // 計算とプレビュー / Calculation and preview
    // =========================================

    /**
     * 比率を向きに合わせて反転する
     * @param {number} ratio - 比率（横 ÷ 縦）
     * @param {boolean} wantPortrait - 縦置きなら true
     * @returns {number} 向きをそろえた比率
     */
    function orientRatio(ratio, wantPortrait) {
        if (wantPortrait && ratio > 1) return 1 / ratio;
        if (!wantPortrait && ratio < 1) return 1 / ratio;
        return ratio;
    }

    /**
     * 固定する辺の長さから、比率に合う幅と高さを求める
     * @param {number} orientedRatio - 向きをそろえた比率
     * @param {boolean} fixByHeight - 高さを固定するなら true
     * @param {number} baseSizePt - 固定する辺の長さ（pt）
     * @returns {{width: number, height: number}} 幅と高さ（pt）
     */
    function computeTargetSize(orientedRatio, fixByHeight, baseSizePt) {
        /* 比率から求める側だけを単位に合わせて丸める / Round only the side derived from the ratio */
        if (fixByHeight) {
            return { width: roundForUnit(baseSizePt * orientedRatio), height: baseSizePt };
        }
        return { width: baseSizePt, height: roundForUnit(baseSizePt / orientedRatio) };
    }

    /**
     * プレビューのアイテムに比率を当てる
     * @param {Object} preview - { items, originalWidths, originalHeights, originalBounds }
     * @param {Object} scaleSettings - readScaleSettings() の戻り値
     * @returns {void}
     */
    function applyRatioToPreview(preview, scaleSettings) {
        var orientedRatio = orientRatio(scaleSettings.ratio, scaleSettings.wantPortrait);
        for (var i = 0; i < preview.items.length; i++) {
            var baseSizePt = scaleSettings.targetSizePt;
            if (baseSizePt === null) {
                baseSizePt = scaleSettings.fixByHeight ? preview.originalHeights[i] : preview.originalWidths[i];
            }
            var targetSize = computeTargetSize(orientedRatio, scaleSettings.fixByHeight, baseSizePt);
            var previewItem = preview.items[i];
            previewItem.width = targetSize.width;
            previewItem.height = targetSize.height;

            /* 基準点が元の位置に戻るように移動 / Move so the reference point returns to where it was */
            var anchorBefore = getAnchorPoint(preview.originalBounds[i], scaleSettings.anchorIndex);
            var anchorAfter = getAnchorPoint(previewItem.geometricBounds, scaleSettings.anchorIndex);
            previewItem.translate(anchorBefore[0] - anchorAfter[0], anchorBefore[1] - anchorAfter[1]);
        }
        app.redraw();
    }

    /**
     * 選択アイテムの複製をプレビュー用に作り、元は隠す
     * @param {Object[]} selectedItems - 選択アイテム
     * @returns {Object} { items, originalWidths, originalHeights, originalBounds }
     */
    function createPreviewFromSelection(selectedItems) {
        var preview = { items: [], originalWidths: [], originalHeights: [], originalBounds: [] };
        for (var i = 0; i < selectedItems.length; i++) {
            var previewCopy = selectedItems[i].duplicate();
            previewCopy.zOrder(ZOrderMethod.BRINGTOFRONT);
            preview.items.push(previewCopy);
            preview.originalWidths.push(selectedItems[i].width);
            preview.originalHeights.push(selectedItems[i].height);
            preview.originalBounds.push(selectedItems[i].geometricBounds);
            selectedItems[i].hidden = true;
        }
        return preview;
    }

    /**
     * 選択がないとき、アクティブなアートボードの中央に長方形を作ってプレビューにする
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} scaleSettings - readScaleSettings() の戻り値
     * @returns {Object} { items, originalWidths, originalHeights, originalBounds }
     */
    function createPreviewRectangle(doc, scaleSettings) {
        var baseSizePt = (scaleSettings.targetSizePt === null) ? FALLBACK_BASE_SIZE_PT : scaleSettings.targetSizePt;
        var targetSize = computeTargetSize(orientRatio(scaleSettings.ratio, scaleSettings.wantPortrait), scaleSettings.fixByHeight, baseSizePt);

        var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; /* [L,T,R,B] */
        var centerX = (artboardRect[0] + artboardRect[2]) / 2;
        var centerY = (artboardRect[1] + artboardRect[3]) / 2;
        var previewRect = doc.pathItems.rectangle(centerY + targetSize.height / 2, centerX - targetSize.width / 2, targetSize.width, targetSize.height);
        previewRect.stroked = false;
        previewRect.filled = true;

        return {
            items: [previewRect],
            originalWidths: [targetSize.width],
            originalHeights: [targetSize.height],
            originalBounds: [previewRect.geometricBounds]
        };
    }

    // =========================================
    // 確定と取り消し / Commit and cancel
    // =========================================

    /**
     * 確定後の仕上げ（ピクセル最適化・アートボード追加）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} item - 仕上げるアイテム
     * @param {Object} dialogControls - createDialog() の戻り値
     * @returns {void}
     */
    function applyFinishingOptions(doc, item, dialogControls) {
        if (dialogControls.alignToPixelCheckbox.value) {
            doc.selection = [item];
            app.executeMenuCommand("Make Pixel Perfect");
        }
        if (dialogControls.addArtboardCheckbox.value) {
            doc.artboards.add(item.visibleBounds);
        }
    }

    /**
     * プレビューの大きさと位置を元のアイテムに移し、プレビューを消す
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} selectedItems - 元の選択アイテム
     * @param {Object} preview - { items, originalWidths, originalHeights, originalBounds }
     * @param {Object} dialogControls - createDialog() の戻り値
     * @returns {void}
     */
    function commitToOriginals(doc, selectedItems, preview, dialogControls) {
        for (var i = 0; i < selectedItems.length; i++) {
            var originalItem = selectedItems[i];
            var previewCopy = preview.items[i];
            originalItem.hidden = false;

            /* 中心基準で拡大縮小し、左上をそろえる / Scale from center, then align the top-left */
            var originalWidth = preview.originalWidths[i];
            var originalHeight = preview.originalHeights[i];
            if (originalWidth > 0 && originalHeight > 0) {
                originalItem.resize(previewCopy.width / originalWidth * 100, previewCopy.height / originalHeight * 100);
            }
            originalItem.position = previewCopy.position;
            previewCopy.remove();

            applyFinishingOptions(doc, originalItem, dialogControls);
        }
        doc.selection = selectedItems;
    }

    /**
     * プレビューを消し、隠した元のアイテムを戻す
     * @param {Object[]} selectedItems - 元の選択アイテム
     * @param {Object} preview - { items, originalWidths, originalHeights, originalBounds }
     * @returns {void}
     */
    function cancelPreview(selectedItems, preview) {
        for (var i = 0; i < preview.items.length; i++) {
            preview.items[i].remove();
        }
        for (var j = 0; j < selectedItems.length; j++) {
            selectedItems[j].hidden = false;
        }
    }

    /**
     * 固定しない辺に、比率から求めた長さを表示する（対象が複数なら空欄）
     * @param {Object} dialogControls - createDialog() の戻り値
     * @param {Object} preview - { items, originalWidths, originalHeights, originalBounds }
     * @param {boolean} fixByHeight - 高さを固定するなら true
     * @returns {void}
     */
    function showComputedSide(dialogControls, preview, fixByHeight) {
        var computedRow = fixByHeight ? dialogControls.widthRow : dialogControls.heightRow;
        var lengthText = "";
        if (preview.items.length === 1) {
            var lengthPt = fixByHeight ? preview.items[0].width : preview.items[0].height;
            lengthText = String(Math.round(lengthPt / RULER_UNIT.pointsPerUnit * 100) / 100);
        }
        /* 表示と、固定する辺に切り替えたときの初期値の両方に使う / Used for display and as the value when this side becomes fixed */
        computedRow.valueText.text = lengthText;
        computedRow.field.text = lengthText;
    }

    /**
     * ダイアログの状態をそろえてプレビューを更新する
     * @param {Object} dialogControls - createDialog() の戻り値
     * @param {Object} preview - { items, originalWidths, originalHeights, originalBounds }
     * @returns {void}
     */
    function refreshPreview(dialogControls, preview) {
        var isCustomRatio = dialogControls.ratioCustomRadio.value;
        dialogControls.customWidthField.enabled = isCustomRatio;
        dialogControls.customHeightField.enabled = isCustomRatio;
        var fixByHeight = dialogControls.basisVerticalRadio.value;
        setSizeRowEditable(dialogControls.widthRow, !fixByHeight);
        setSizeRowEditable(dialogControls.heightRow, fixByHeight);
        applyRatioToPreview(preview, readScaleSettings(dialogControls));
        showComputedSide(dialogControls, preview, fixByHeight);
    }

    /**
     * 設定を変えるたびにプレビューを更新するよう、各コントロールにハンドラーを付ける
     * @param {Object} dialogControls - createDialog() の戻り値
     * @param {Function} onSettingsChange - 呼び出す関数
     * @returns {void}
     */
    function bindSettingsHandlers(dialogControls, onSettingsChange) {
        var clickableControls = [
            dialogControls.ratio16x9Radio, dialogControls.ratioSquareRadio, dialogControls.ratioA4Radio, dialogControls.ratioCustomRadio,
            dialogControls.basisHorizontalRadio, dialogControls.basisVerticalRadio
        ];
        for (var i = 0; i < clickableControls.length; i++) {
            clickableControls[i].onClick = onSettingsChange;
        }
        var numberFields = [
            dialogControls.customWidthField, dialogControls.customHeightField,
            dialogControls.widthRow.field, dialogControls.heightRow.field
        ];
        for (var j = 0; j < numberFields.length; j++) {
            numberFields[j].onChanging = onSettingsChange;
        }
        dialogControls.anchorWidget.onAnchorChange = onSettingsChange;
        dialogControls.orientationGroup.onOrientationChange = onSettingsChange;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * メイン処理
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        var hasSelection = (selectedItems && selectedItems.length > 0);
        if (!hasSelection) selectedItems = [];

        var dialogControls = createDialog();

        var preview;
        if (hasSelection) {
            preview = createPreviewFromSelection(selectedItems);
        } else {
            /* 選択なしのときは幅の欄に既定値（px:1000 / mm:100）/ Default width when nothing is selected */
            dialogControls.widthRow.field.text = DEFAULT_SIZE_TEXT_BY_UNIT[RULER_UNIT.code] || "";
            preview = createPreviewRectangle(doc, readScaleSettings(dialogControls));
        }

        /* 設定が変わるたびにプレビューを更新 / Refresh the preview on every change */
        bindSettingsHandlers(dialogControls, function () { refreshPreview(dialogControls, preview); });
        refreshPreview(dialogControls, preview);

        if (dialogControls.dialog.show() !== 1) {
            cancelPreview(selectedItems, preview);
            return;
        }

        if (hasSelection) {
            commitToOriginals(doc, selectedItems, preview, dialogControls);
        } else {
            /* 作った長方形をそのまま残す / Keep the preview rectangle as the result */
            applyFinishingOptions(doc, preview.items[0], dialogControls);
            doc.selection = [preview.items[0]];
        }
        app.redraw();
    }

    main();

})();
