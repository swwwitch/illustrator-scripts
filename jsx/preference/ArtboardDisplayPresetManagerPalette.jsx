#target illustrator
#targetengine "ArtboardDisplayPresetManagerPalette"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボード関連のIllustrator環境設定を、常駐パレットでまとめて切り替えます。
アートボードのサイズ確認とリサイズ、名前表示と枠線、プリセット、カンバスカラーの切り替えに対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardDisplayPresetManagerPalette.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n9eba8ab03170

### Overview

A persistent palette for switching the artboard-related Illustrator preferences.
It covers checking and resizing artboards, name and border display, presets, and the canvas color.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardDisplayPresetManagerPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArtboardDisplayPresetManagerPalette"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-26";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardDisplayPresetManagerPalette.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardDisplayPresetManagerPalette.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n9eba8ab03170"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 枠線カラーの選択肢（ドロップダウンの並び順。colorKey は LABELS.dropdown.borderColor のキー）
       Border color choices (dropdown order; colorKey is a key of LABELS.dropdown.borderColor) */
    var BORDER_COLOR_PRESETS = [
        { colorKey: "lightBlue",  r: 0.29, g: 0.52, b: 1.0 },
        { colorKey: "lightRed",   r: 1.0,  g: 0.29, b: 0.29 },
        { colorKey: "green",      r: 0.0,  g: 0.65, b: 0.31 },
        { colorKey: "mediumBlue", r: 0.0,  g: 0.45, b: 0.78 },
        { colorKey: "magenta",    r: 1.0,  g: 0.0,  b: 1.0 },
        { colorKey: "cyan",       r: 0.0,  g: 1.0,  b: 1.0 },
        { colorKey: "lightGray",  r: 0.65, g: 0.65, b: 0.65 },
        { colorKey: "black",      r: 0.0,  g: 0.0,  b: 0.0 },
        { colorKey: "yellow",     r: 1.0,  g: 1.0,  b: 0.0 }
    ];

    /* 枠線の太さの選択肢（1〜4）/ Border width choices (1-4) */
    var BORDER_WIDTH_CHOICES = [1, 2, 3, 4];

    /* 表示プリセット（適用・判定の両方で使う単一の真実）/ Display presets (single source for both apply and detect) */
    /* "default" は ES3 予約語のため引用符付きキー / "default" is an ES3 reserved word, so quote the key */
    /* printBleed は現在未使用（PRINT_BLEED_WIDGET を参照）/ printBleed is currently unused (see the PRINT_BLEED_WIDGET note) */
    var PRESET_KEYS = ["default", "emphasis", "light"];
    var PRESETS = {
        "default": { showName: true,  colorKey: "black",     borderWidth: 1, printBleed: true,  moveLocked: false },
        emphasis:  { showName: false, colorKey: "lightRed",  borderWidth: 3, printBleed: false, moveLocked: true },
        light:     { showName: false, colorKey: "lightGray", borderWidth: 1, printBleed: false, moveLocked: true }
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;               /* パネル内の要素間隔 / panel spacing */
    var SIZE_LABEL_WIDTH = 48;             /* 幅・高さラベルの幅（px）/ width of the width/height labels */

    /* 9軸ウィジェット / 9-axis anchor widget */
    var ANCHOR_WIDGET_SIZE   = 66;         /* ウィジェットの一辺（px）/ widget size */
    var ANCHOR_CELL_SIZE     = 9;          /* □1個のサイズ / size of one anchor square */
    var ANCHOR_CELL_GAP      = 7.5;        /* □どうしの間隔 / gap between anchor squares */
    var DEFAULT_ANCHOR_INDEX = 0;          /* 既定の基準点（0=左上〜8=右下の行優先）/ default anchor (row-major, 0=top-left..8=bottom-right) */

    /**
     * ウィンドウの共通設定を適用する
     * @param {Window} win - 対象ウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定を適用する
     * @param {Panel} panel - 対象パネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 行グループの共通設定を適用する（ボタン列・ラベル＋入力欄など）
     * statictext は edittext より低く ScriptUI(mac) では上寄せになりやすいため、天地中央を明示する
     * @param {Group} rowGroup - 対象グループ
     * @param {string} [alignment] - グループ自身の整列（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, alignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.alignment = alignment || "left";
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ）
     * @param {Button} button - 対象ボタン
     * @param {number} px - 詰める高さ（px）
     * @returns {void}
     */
    function trimButtonHeight(button, px) {
        button.size = [button.size.width, button.size.height - px];
    }

    /**
     * 数値入力欄を↑↓キーで増減する（shift で±10、option で±0.1）
     * @param {EditText} editText - 対象の入力欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName != "Up" && event.keyName != "Down") return;
            var currentValue = Number(editText.text);
            if (isNaN(currentValue)) return;
            var keyboard = ScriptUI.environment.keyboardState;
            var delta = (event.keyName == "Up") ? 1 : -1;

            if (keyboard.shiftKey) {
                /* Shift：10の倍数にスナップ / Snap to multiples of 10 */
                currentValue = Math.round(currentValue / 10) * 10 + delta * 10;
                if (currentValue < 0) currentValue = 0;
            } else if (keyboard.altKey) {
                /* Option：±0.1（小数第1位に丸め）/ ±0.1, rounded to one decimal */
                currentValue = Math.round((currentValue + delta * 0.1) * 10) / 10;
            } else {
                currentValue += delta;
                if (currentValue < 0) currentValue = 0;
            }

            event.preventDefault();
            editText.text = currentValue;
            /* プログラム変更は onChanging を発火しないため明示的に呼ぶ / Fire onChanging manually */
            if (typeof editText.onChanging === "function") editText.onChanging();
        });
    }

    /**
     * ↑↓キーを離したときに onChange を呼ぶ（押している間のキーリピートでは確定しない）
     * Illustrator にはタイマーが無いため、キーを離すまでを待ち時間の代わりにする
     * @param {EditText} editText - 対象の入力欄
     * @returns {void}
     */
    function commitOnArrowKeyUp(editText) {
        editText.addEventListener("keyup", function (event) {
            if (event.keyName != "Up" && event.keyName != "Down") return;
            if (typeof editText.onChange === "function") editText.onChange();
        });
    }

    // =========================================
    // 9軸ウィジェット / 9-axis anchor widget
    // =========================================

    /* 枠線は常時薄いグレー、選択セルの塗りは initAnchorColors() で UI 明暗に合わせる / Border is always light gray; selected fill follows the UI brightness */
    var ANCHOR_LINE_COLOR = [0.6, 0.6, 0.6, 1];
    var ANCHOR_SELECTED_FILL = [0.4, 0.4, 0.4, 1];
    var ANCHOR_DISABLED_COLOR = [0.5, 0.5, 0.5, 0.4];

    /* 中央(4)を除く外周の□どうしをつなぐケイ線の組み合わせ / Pairs of outer squares (center 4 excluded) joined by rules */
    var ANCHOR_CONNECTIONS = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];

    /**
     * UI の明暗から選択セルの塗りを決める（ライトは濃いグレー、ダークは明るいグレー）
     * @returns {void}
     */
    function initAnchorColors() {
        var isLightUI = readPref("Real", "uiBrightness", 0) > 0.5;
        ANCHOR_SELECTED_FILL = isLightUI ? [0.4, 0.4, 0.4, 1] : [0.8, 0.8, 0.8, 1];
    }

    /**
     * 正方形のパスを作る（塗り／線は呼び出し側で行う）
     * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
     * @param {number} x - 左端
     * @param {number} y - 上端
     * @param {number} size - 一辺の長さ
     * @returns {void}
     */
    function squarePath(graphics, x, y, size) {
        graphics.newPath();
        graphics.moveTo(x, y);
        graphics.lineTo(x + size, y);
        graphics.lineTo(x + size, y + size);
        graphics.lineTo(x, y + size);
        graphics.closePath();
    }

    /**
     * 9軸ウィジェットを描画する（外周の□をケイ線でつなぐ・中央は独立、選択セルだけ塗る）
     * @param {Button} widget - 対象のウィジェット
     * @returns {void}
     */
    function drawAnchorWidget(widget) {
        var graphics = widget.graphics;
        var width = widget.size[0];
        var height = widget.size[1];

        /* コントロール地色で塗って透過に見せる / Paint the control's own background so it looks transparent */
        try {
            graphics.newPath();
            graphics.rectPath(0, 0, width, height);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) {}

        var lineColor = widget.enabled ? ANCHOR_LINE_COLOR : ANCHOR_DISABLED_COLOR;
        var fillColor = widget.enabled ? ANCHOR_SELECTED_FILL : ANCHOR_DISABLED_COLOR;
        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1);
        var cellStep = ANCHOR_CELL_SIZE + ANCHOR_CELL_GAP;
        var gridSize = ANCHOR_CELL_SIZE * 3 + ANCHOR_CELL_GAP * 2;
        var originX = Math.round((width - gridSize) / 2);
        var originY = Math.round((height - gridSize) / 2);

        var cellPositions = [];
        for (var index = 0; index < 9; index++) {
            cellPositions.push([originX + (index % 3) * cellStep, originY + Math.floor(index / 3) * cellStep]);
        }

        for (var i = 0; i < ANCHOR_CONNECTIONS.length; i++) {
            var cellA = cellPositions[ANCHOR_CONNECTIONS[i][0]];
            var cellB = cellPositions[ANCHOR_CONNECTIONS[i][1]];
            graphics.newPath();
            if (ANCHOR_CONNECTIONS[i][1] - ANCHOR_CONNECTIONS[i][0] === 1) {
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

        for (var cellIndex = 0; cellIndex < cellPositions.length; cellIndex++) {
            /* 枠を上に描くので塗りを先に行う / Fill first so the border draws on top */
            if (cellIndex === widget.selectedAnchorIndex) {
                squarePath(graphics, cellPositions[cellIndex][0], cellPositions[cellIndex][1], ANCHOR_CELL_SIZE);
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, fillColor));
            }
            squarePath(graphics, cellPositions[cellIndex][0], cellPositions[cellIndex][1], ANCHOR_CELL_SIZE);
            graphics.strokePath(linePen);
        }
    }

    /**
     * 9軸（3×3）の基準点ウィジェットを生成する（クリックしたセルを基準点にする）
     * @param {Group} parentGroup - 追加先のグループ
     * @returns {Button} 生成したウィジェット（選択は selectedAnchorIndex に保持）
     */
    function addAnchorWidget(parentGroup) {
        var widget = parentGroup.add("button", undefined, "");
        widget.helpTip = getLabel("tooltip.anchor");
        widget.preferredSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        widget.minimumSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        widget.maximumSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        widget.selectedAnchorIndex = DEFAULT_ANCHOR_INDEX;
        widget.onDraw = function () { drawAnchorWidget(this); };
        /* クリック座標（コントロール基準）を3分割してセルを判定 / Hit-test by splitting the control-relative click into thirds */
        widget.addEventListener("mousedown", function (event) {
            var col = Math.min(2, Math.max(0, Math.floor(event.clientX / (widget.size[0] / 3))));
            var row = Math.min(2, Math.max(0, Math.floor(event.clientY / (widget.size[1] / 3))));
            widget.selectedAnchorIndex = row * 3 + col;
            try { widget.notify("onDraw"); } catch (e) {}
        });
        return widget;
    }

    // =========================================
    // 常駐エンジン / Persistent engine
    // =========================================

    /* 常駐エンジンにパレット参照を保持するキー / Key holding the palette reference in the persistent engine */
    var PALETTE_GLOBAL_KEY = "__artboardDisplayPresetPalette";

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

    /**
     * pt 値を指定単位の表示文字列へ変換する（小数2桁で丸め）
     * @param {number} pointValue - pt 値
     * @param {object} unitInfo - getUnitInfo() の戻り値
     * @returns {string} 表示用文字列（数値でない場合は空文字）
     */
    function pointToUnitText(pointValue, unitInfo) {
        if (isNaN(pointValue)) return "";
        return String(Math.round(pointValue / unitInfo.pointsPerUnit * 100) / 100);
    }

    // =========================================
    // 環境設定アクセス / Preferences access
    // =========================================

    var appPreferences = app.preferences;

    /* -----------------------------------------
       PRINT_BLEED_WIDGET：「裁ち落としを印刷」生成AIボタンの表示切り替えについて

       enablePrintBleedWidget は setBooleanPreference で確実に書き換えられ、環境設定ダイアログの
       表示にも反映されるが、カンバス上のウィジェットは再評価されない。app.redraw()、ズーム、
       ツール切り替え、プレビュー／アウトライン切り替え、アートボード再設定、ドキュメント切り替えの
       いずれでも反映せず、スクリプトから即時反映させる手段が見つからなかった（Illustrator が
       起動時か環境設定ダイアログの適用時にしか読まないと思われる）。

       将来のバージョンで挙動が戻る可能性があるため、関連コードは削除せずコメントアウトで保留する。
       復活させるときは、この定数・LABELS の checkbox.showPrintBleedAI・buildOptionsPanel・
       applyPrintBleedWidgetSetting・applyOptionSettings・applyPresetToUI・findMatchingPresetKey・
       reflectPreferences・wirePaletteEvents のコメントアウトを戻すこと。

       The preference is written reliably, but Illustrator never re-evaluates the canvas widget,
       and no scripted refresh triggers it. The item is parked (commented out) rather than removed.
       ----------------------------------------- */
    // var PRINT_BLEED_WIDGET_KEY = "enablePrintBleedWidget";

    /**
     * 環境設定を型指定で読み取る
     * 未登録のキーや型の食い違いで例外になることがあるため try で受ける
     * @param {string} valueType - "Real" | "Boolean" | "Integer"
     * @param {string} key - 環境設定キー
     * @param {*} fallback - 取得に失敗したときの値
     * @returns {*} 取得値（失敗時は fallback）
     */
    function readPref(valueType, key, fallback) {
        try { return appPreferences["get" + valueType + "Preference"](key); } catch (e) { return fallback; }
    }

    /**
     * メニューコマンドを実行する
     * ドキュメントが無いときなど、実行できない状況の例外は無視する
     * @param {string} command - メニューコマンド名
     * @returns {void}
     */
    function runMenuCommand(command) {
        try { app.executeMenuCommand(command); } catch (e) {}
    }

    /**
     * アートボード表示を強制的に再描画する（環境設定の変更は自動では反映されないため）
     * @returns {void}
     */
    function refreshArtboardDisplay() {
        runMenuCommand("zoomout");
        runMenuCommand("zoomin");
    }

    // =========================================
    // アートボード枠線カラー / Artboard border color
    // =========================================

    /* colorKey が見つからないときに使うブラックの index / Black index used when a colorKey is not found */
    var BORDER_COLOR_BLACK_INDEX = 7;

    /**
     * ドロップダウン用のカラー名配列を生成する
     * @returns {string[]} 現在言語のカラー名の配列
     */
    function buildBorderColorNames() {
        var colorNames = [];
        for (var i = 0; i < BORDER_COLOR_PRESETS.length; i++) {
            colorNames.push(getLabel("dropdown.borderColor." + BORDER_COLOR_PRESETS[i].colorKey));
        }
        return colorNames;
    }

    /**
     * 指定 RGB にもっとも近いカラープリセットの index を返す
     * @param {number} red - 赤成分（0〜1）
     * @param {number} green - 緑成分（0〜1）
     * @param {number} blue - 青成分（0〜1）
     * @returns {number} もっとも近いプリセットの index
     */
    function findClosestBorderColorIndex(red, green, blue) {
        var closestIndex = 0;
        var closestDistance = Infinity;
        for (var i = 0; i < BORDER_COLOR_PRESETS.length; i++) {
            var colorPreset = BORDER_COLOR_PRESETS[i];
            var distance = Math.abs(colorPreset.r - red) + Math.abs(colorPreset.g - green) + Math.abs(colorPreset.b - blue);
            if (distance < closestDistance) {
                closestDistance = distance;
                closestIndex = i;
            }
        }
        return closestIndex;
    }

    /**
     * colorKey からカラープリセットの index を取得する
     * @param {string} colorKey - BORDER_COLOR_PRESETS の colorKey（例: "lightRed"）
     * @returns {number} プリセットの index（見つからない場合はブラック）
     */
    function findBorderColorIndexByKey(colorKey) {
        for (var i = 0; i < BORDER_COLOR_PRESETS.length; i++) {
            if (BORDER_COLOR_PRESETS[i].colorKey === colorKey) return i;
        }
        return BORDER_COLOR_BLACK_INDEX;
    }

    // =========================================
    // アートボード情報（BridgeTalk 委譲）/ Artboard info (BridgeTalk delegation)
    // 常駐パレットのイベントハンドラ内では DOM 接続を失うため、DOM の読み書きはメインエンジンへ委譲する
    // Inside persistent-palette event handlers the DOM connection is lost, so all DOM access is delegated to the main engine
    // =========================================

    /* フィールド区切り（アートボード名に現れにくい ASCII 文字列／エスケープ不要）/ Field separator (ASCII, unlikely in names, no escaping) */
    var ARTBOARD_FIELD_SEPARATOR = "<|>";

    /**
     * メインエンジンへコードを委譲し、結果文字列をコールバックへ渡す
     * @param {string} bodyCode - メインエンジンで実行するコード
     * @param {function} [onResult] - 結果文字列を受け取るコールバック（失敗時は空文字）
     * @returns {void}
     */
    function delegateToMainEngine(bodyCode, onResult) {
        if (typeof onResult !== "function") onResult = function () {};
        var bridge = new BridgeTalk();
        bridge.target = "illustrator"; /* #targetengine 指定なし＝メインエンジン / no engine = main engine */
        bridge.body = bodyCode;
        bridge.onResult = function (response) { onResult(String(response.body)); };
        bridge.onError = function () { onResult(""); };
        bridge.send();
    }

    /**
     * メインエンジンで実行するコード本体を組み立てる（read／丸め／リサイズ後に最新情報を返す）
     * @param {string} operation - "read" | "round" | "resize"
     * @param {number} [widthPoint] - resize 時の幅（pt）
     * @param {number} [heightPoint] - resize 時の高さ（pt）
     * @param {number} [anchorIndex] - resize 時の基準点（0=左上〜8=右下の行優先）
     * @returns {string} メインエンジンへ送るコード
     */
    function buildArtboardBridgeCode(operation, widthPoint, heightPoint, anchorIndex) {
        var mutationCode = "";
        if (operation === "round") {
            mutationCode = "artboard.artboardRect=[Math.round(rect[0]),Math.round(rect[1]),Math.round(rect[2]),Math.round(rect[3])];rect=artboard.artboardRect;";
        } else if (operation === "resize") {
            /* 負数連結による '--' 構文エラーを避けるため括弧で囲む / Wrap in parens to avoid '--' from negative numbers */
            /* 基準点の側に寄せる：増減分に 0／0.5／1 を掛けて左上をずらす / Shift the top-left by 0, 0.5 or 1 of the size change toward the anchor */
            var anchorColumnRatio = (anchorIndex % 3) / 2;
            var anchorRowRatio = Math.floor(anchorIndex / 3) / 2;
            mutationCode = "var w=" + Number(widthPoint) + ",h=" + Number(heightPoint) + ";" +
                "var left=rect[0]+(rect[2]-rect[0]-w)*" + anchorColumnRatio + ";" +
                "var top=rect[1]-(rect[1]-rect[3]-h)*" + anchorRowRatio + ";" +
                "artboard.artboardRect=[left,top,left+w,top-h];rect=artboard.artboardRect;";
        }
        var separatorCode = "+\"" + ARTBOARD_FIELD_SEPARATOR + "\"+";
        return "" +
            "(function(){try{" +
            "var doc=app.activeDocument;" +
            "var artboardIndex=doc.artboards.getActiveArtboardIndex();" +
            "var artboard=doc.artboards[artboardIndex];var rect=artboard.artboardRect;" +
            mutationCode +
            "return (artboardIndex+1)" + separatorCode + "artboard.name" + separatorCode + "(rect[2]-rect[0])" + separatorCode + "(rect[1]-rect[3]);" +
            "}catch(e){return \"\";}})();";
    }

    // =========================================
    // UI構築 / UI construction
    // =========================================

    /**
     * ラベル・入力欄・単位表示を並べたサイズ入力欄を作る
     * @param {Group} parentRow - 追加先の行グループ
     * @param {string} labelPath - ラベルの LABELS キー
     * @param {string} unitLabel - 単位ラベル（例: "mm"）
     * @returns {object} { input: EditText, unitText: StaticText }
     */
    function addSizeField(parentRow, labelPath, unitLabel) {
        var sizeFieldGroup = parentRow.add("group");
        setupRow(sizeFieldGroup, "left", 4);
        var sizeLabel = sizeFieldGroup.add("statictext", undefined, labelText(labelPath));
        sizeLabel.preferredSize.width = SIZE_LABEL_WIDTH;
        sizeLabel.justify = "right";
        var sizeInput = sizeFieldGroup.add("edittext", undefined, "");
        sizeInput.helpTip = getLabel("tooltip.sizeField");
        sizeInput.characters = 5;
        changeValueByArrowKey(sizeInput);
        commitOnArrowKeyUp(sizeInput);
        var unitText = sizeFieldGroup.add("statictext", undefined, unitLabel);
        return { input: sizeInput, unitText: unitText };
    }

    /**
     * ヘルプ付きのボタンを追加する
     * @param {Group} parentRow - 追加先の行グループ
     * @param {string} labelKey - LABELS.button / LABELS.tooltip 共通のキー
     * @param {Array|string} alignment - ボタンの整列
     * @returns {Button} 追加したボタン
     */
    function addButtonWithTip(parentRow, labelKey, alignment) {
        var button = parentRow.add("button", undefined, getLabel("button." + labelKey));
        button.helpTip = getLabel("tooltip." + labelKey);
        button.alignment = alignment;
        return button;
    }

    /**
     * 「現在のアートボード」パネルを構築する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @returns {object} パネル内のコントロール
     */
    function buildCurrentArtboardPanel(parentWindow) {
        var currentArtboardPanel = parentWindow.add("panel", undefined, getLabel("panel.currentArtboard"));
        setupPanel(currentArtboardPanel, 8);

        /* 番号・名前（左右中央）/ Number and name (centered) */
        var artboardInfoRow = currentArtboardPanel.add("group");
        setupRow(artboardInfoRow, "fill");
        artboardInfoRow.alignChildren = ["center", "center"];
        var artboardInfoText = artboardInfoRow.add("statictext", undefined, "—");
        artboardInfoText.characters = 28;
        artboardInfoText.justify = "center";

        /* 幅・高さ（縦並び）＋基準点の9軸 / Width and height (stacked) + 9-axis anchor */
        var sizeRow = currentArtboardPanel.add("group");
        setupRow(sizeRow, "left", 16);
        var sizeColumn = sizeRow.add("group");
        sizeColumn.orientation = "column";
        sizeColumn.alignChildren = ["left", "center"];
        sizeColumn.alignment = "left";
        sizeColumn.spacing = 6;
        var rulerUnitLabel = getUnitInfo().label;
        var widthField = addSizeField(sizeColumn, "fieldLabel.width", rulerUnitLabel);
        var heightField = addSizeField(sizeColumn, "fieldLabel.height", rulerUnitLabel);
        var anchorWidget = addAnchorWidget(sizeRow);

        /* ボタン行（パネル幅いっぱいには広げない）/ Button row (do not stretch to the panel width) */
        var artboardButtonRow = currentArtboardPanel.add("group");
        setupRow(artboardButtonRow, "left");

        return {
            artboardInfoText: artboardInfoText,
            widthInput: widthField.input,
            heightInput: heightField.input,
            widthUnitText: widthField.unitText,
            heightUnitText: heightField.unitText,
            anchorWidget: anchorWidget,
            optimizePixelGridButton: addButtonWithTip(artboardButtonRow, "optimizePixelGrid", "left")
        };
    }

    /**
     * 同じヘルプを付けたラジオボタンを並べる
     * @param {Group} parentRow - 追加先の行グループ
     * @param {string[]} radioTexts - ラジオボタンの文言
     * @param {string} tooltipText - 共通のヘルプ
     * @returns {RadioButton[]} 追加したラジオボタン
     */
    function addRadioRow(parentRow, radioTexts, tooltipText) {
        var radios = [];
        for (var i = 0; i < radioTexts.length; i++) {
            var radio = parentRow.add("radiobutton", undefined, radioTexts[i]);
            radio.helpTip = tooltipText;
            radios.push(radio);
        }
        return radios;
    }

    /**
     * 「アートボード名と枠線」パネルを構築する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @returns {object} パネル内のコントロール
     */
    function buildArtboardDisplayPanel(parentWindow) {
        var artboardDisplayPanel = parentWindow.add("panel", undefined, getLabel("panel.artboardDisplay"));
        setupPanel(artboardDisplayPanel);

        var showNameCheckbox = artboardDisplayPanel.add("checkbox", undefined, getLabel("checkbox.showArtboardName"));
        showNameCheckbox.helpTip = getLabel("tooltip.showArtboardName");

        /* 枠線サブパネル / Border sub-panel */
        var borderPanel = artboardDisplayPanel.add("panel", undefined, getLabel("panel.artboardBorder"));
        setupPanel(borderPanel, 8);

        var borderColorRow = borderPanel.add("group");
        setupRow(borderColorRow, "left");
        borderColorRow.add("statictext", undefined, labelText("fieldLabel.borderColor"));
        var borderColorList = borderColorRow.add("dropdownlist", undefined, buildBorderColorNames());
        borderColorList.helpTip = getLabel("tooltip.borderColor");

        var borderWidthRow = borderPanel.add("group");
        setupRow(borderWidthRow, "left");
        borderWidthRow.add("statictext", undefined, labelText("fieldLabel.borderWidth"));
        var borderWidthTexts = [];
        for (var i = 0; i < BORDER_WIDTH_CHOICES.length; i++) {
            borderWidthTexts.push(String(BORDER_WIDTH_CHOICES[i]));
        }

        /* プリセット（パネル最下部・左右中央）/ Presets (bottom of panel, centered) */
        var presetRow = artboardDisplayPanel.add("group");
        setupRow(presetRow, "center");
        var presetTexts = [];
        for (var j = 0; j < PRESET_KEYS.length; j++) {
            presetTexts.push(getLabel("radio.preset." + PRESET_KEYS[j]));
        }

        return {
            showNameCheckbox: showNameCheckbox,
            borderColorList: borderColorList,
            borderWidthRadios: addRadioRow(borderWidthRow, borderWidthTexts, getLabel("tooltip.borderWidth")),
            presetRadios: addRadioRow(presetRow, presetTexts, getLabel("tooltip.preset"))
        };
    }

    /**
     * 「オプション」パネルを構築する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @returns {object} パネル内のコントロール
     */
    function buildOptionsPanel(parentWindow) {
        var optionsPanel = parentWindow.add("panel", undefined, getLabel("panel.options"));
        setupPanel(optionsPanel, 8);

        /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
        // var printBleedCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.showPrintBleedAI"));

        var moveLockedHiddenCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.moveLockedHidden"));
        moveLockedHiddenCheckbox.helpTip = getLabel("tooltip.moveLockedHidden");

        return {
            // printBleedCheckbox: printBleedCheckbox,
            moveLockedHiddenCheckbox: moveLockedHiddenCheckbox
        };
    }

    /**
     * 下部のボタン行（カンバスカラー／ビデオ定規）を構築する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @returns {object} 行内のボタン
     */
    function buildFooterRow(parentWindow) {
        var footerRow = parentWindow.add("group");
        setupRow(footerRow, "fill");

        var canvasColorButton = addButtonWithTip(footerRow, "canvasColor", ["left", "center"]);
        var footerSpacer = footerRow.add("group");
        footerSpacer.alignment = ["fill", "fill"];
        var videoRulerButton = addButtonWithTip(footerRow, "videoRuler", ["right", "center"]);

        return { canvasColorButton: canvasColorButton, videoRulerButton: videoRulerButton };
    }

    /**
     * パレット全体を構築する
     * @returns {object} ウィンドウと全コントロールをまとめたオブジェクト
     */
    function buildPalette() {
        var paletteWindow = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(paletteWindow);
        initAnchorColors();

        var paletteUI = { paletteWindow: paletteWindow };
        var sectionControls = [
            buildCurrentArtboardPanel(paletteWindow),
            buildArtboardDisplayPanel(paletteWindow),
            buildOptionsPanel(paletteWindow),
            buildFooterRow(paletteWindow)
        ];
        for (var i = 0; i < sectionControls.length; i++) {
            for (var controlName in sectionControls[i]) {
                if (sectionControls[i].hasOwnProperty(controlName)) paletteUI[controlName] = sectionControls[i][controlName];
            }
        }
        return paletteUI;
    }

    // =========================================
    // 即時反映：UI → 環境設定 / Immediate apply: UI to preferences
    // =========================================

    /**
     * 選択中の枠線の太さ（1〜4）を取得する
     * @param {object} paletteUI - パレットのコントロール一式
     * @returns {number} 選択中の太さ
     */
    function getSelectedBorderWidth(paletteUI) {
        for (var i = 0; i < paletteUI.borderWidthRadios.length; i++) {
            if (paletteUI.borderWidthRadios[i].value) return BORDER_WIDTH_CHOICES[i];
        }
        return BORDER_WIDTH_CHOICES[0];
    }

    /**
     * アートボード名・枠線カラー・枠線の太さを環境設定へ反映する
     * @param {object} paletteUI - パレットのコントロール一式
     * @returns {void}
     */
    function applyArtboardDisplaySettings(paletteUI) {
        appPreferences.setBooleanPreference("showArtboardLabelOnCanvas", paletteUI.showNameCheckbox.value);
        var colorIndex = paletteUI.borderColorList.selection ? paletteUI.borderColorList.selection.index : BORDER_COLOR_BLACK_INDEX;
        var borderColor = BORDER_COLOR_PRESETS[colorIndex];
        appPreferences.setRealPreference("ArtboardBBColorRed", borderColor.r);
        appPreferences.setRealPreference("ArtboardBBColorGreen", borderColor.g);
        appPreferences.setRealPreference("ArtboardBBColorBlue", borderColor.b);
        appPreferences.setRealPreference("ArtboardBBWidth", getSelectedBorderWidth(paletteUI));
        refreshArtboardDisplay();
    }

    // PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note
    // 「裁ち落としを印刷」生成AIボタンの表示設定を、書き込み・再描画・読み戻しまで
    // メインエンジンへ委譲する（書き込めなかった場合はチェックを実際の値へ戻す）
    // function applyPrintBleedWidgetSetting(paletteUI) {
    //     var enabled = paletteUI.printBleedCheckbox.value;
    //     delegateToMainEngine("" +
    //         "(function(){" +
    //         "app.preferences.setBooleanPreference('" + PRINT_BLEED_WIDGET_KEY + "'," + (enabled ? "true" : "false") + ");" +
    //         "try{if(app.documents.length){app.activeDocument.activate();app.redraw();" +
    //         "app.executeMenuCommand('zoomout');app.executeMenuCommand('zoomin');}}catch(e){}" +
    //         "return String(app.preferences.getBooleanPreference('" + PRINT_BLEED_WIDGET_KEY + "'));" +
    //         "})();",
    //         function (result) {
    //             if (result === "true" || result === "false") {
    //                 paletteUI.printBleedCheckbox.value = (result === "true");
    //             }
    //         });
    // }

    /**
     * オプション（ロックまたは非表示オブジェクトを一緒に移動）を環境設定へ反映する
     * @param {object} paletteUI - パレットのコントロール一式
     * @returns {void}
     */
    function applyOptionSettings(paletteUI) {
        // applyPrintBleedWidgetSetting(paletteUI); /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
        appPreferences.setBooleanPreference("moveLockedAndHiddenArt", paletteUI.moveLockedHiddenCheckbox.value);
    }

    // =========================================
    // 値反映：環境設定 → UI / Reflect preferences into the UI
    // =========================================

    /**
     * 枠線の太さのラジオボタンを選択する
     * @param {object} paletteUI - パレットのコントロール一式
     * @param {number} borderWidth - 選択する太さ（1〜4）
     * @returns {void}
     */
    function selectBorderWidthRadio(paletteUI, borderWidth) {
        for (var i = 0; i < paletteUI.borderWidthRadios.length; i++) {
            paletteUI.borderWidthRadios[i].value = (BORDER_WIDTH_CHOICES[i] === borderWidth);
        }
    }

    /**
     * プリセットの内容を UI へ適用する
     * @param {object} paletteUI - パレットのコントロール一式
     * @param {string} presetKey - PRESETS のキー
     * @returns {void}
     */
    function applyPresetToUI(paletteUI, presetKey) {
        var preset = PRESETS[presetKey];
        if (!preset) return;
        paletteUI.showNameCheckbox.value = preset.showName;
        paletteUI.borderColorList.selection = findBorderColorIndexByKey(preset.colorKey);
        selectBorderWidthRadio(paletteUI, preset.borderWidth);
        // paletteUI.printBleedCheckbox.value = preset.printBleed; /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
        paletteUI.moveLockedHiddenCheckbox.value = preset.moveLocked;
    }

    /**
     * 現在の UI 状態に一致するプリセットキーを返す
     * @param {object} paletteUI - パレットのコントロール一式
     * @param {number} colorIndex - 現在のカラープリセット index
     * @param {number} borderWidth - 現在の枠線の太さ
     * @returns {string|null} 一致するプリセットキー（なければ null）
     */
    function findMatchingPresetKey(paletteUI, colorIndex, borderWidth) {
        for (var i = 0; i < PRESET_KEYS.length; i++) {
            var preset = PRESETS[PRESET_KEYS[i]];
            if (paletteUI.showNameCheckbox.value === preset.showName &&
                colorIndex === findBorderColorIndexByKey(preset.colorKey) &&
                borderWidth === preset.borderWidth &&
                /* paletteUI.printBleedCheckbox.value === preset.printBleed && */ /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
                paletteUI.moveLockedHiddenCheckbox.value === preset.moveLocked) {
                return PRESET_KEYS[i];
            }
        }
        return null;
    }

    /**
     * 現在の環境設定を UI へ反映し、一致するプリセットのラジオボタンを選択する
     * @param {object} paletteUI - パレットのコントロール一式
     * @returns {void}
     */
    function reflectPreferences(paletteUI) {
        paletteUI.showNameCheckbox.value = !!readPref("Boolean", "showArtboardLabelOnCanvas", false);

        var colorIndex = findClosestBorderColorIndex(
            readPref("Real", "ArtboardBBColorRed", 0.0),
            readPref("Real", "ArtboardBBColorGreen", 0.0),
            readPref("Real", "ArtboardBBColorBlue", 0.0)
        );
        paletteUI.borderColorList.selection = colorIndex;

        /* 1〜4 の範囲に収める / Clamp to 1-4 */
        var borderWidth = Math.max(1, Math.min(BORDER_WIDTH_CHOICES.length, Math.round(readPref("Real", "ArtboardBBWidth", 1.0))));
        selectBorderWidthRadio(paletteUI, borderWidth);

        // paletteUI.printBleedCheckbox.value = !!readPref("Boolean", PRINT_BLEED_WIDGET_KEY, false); /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
        paletteUI.moveLockedHiddenCheckbox.value = !!readPref("Boolean", "moveLockedAndHiddenArt", false);

        var matchedPresetKey = findMatchingPresetKey(paletteUI, colorIndex, borderWidth);
        for (var i = 0; i < PRESET_KEYS.length; i++) {
            paletteUI.presetRadios[i].value = (PRESET_KEYS[i] === matchedPresetKey);
        }
    }

    // =========================================
    // 現在のアートボード / Current artboard
    // =========================================

    /**
     * 委譲結果（番号<|>名前<|>幅pt<|>高さpt）を UI へ反映する
     * @param {object} paletteUI - パレットのコントロール一式
     * @param {string} result - メインエンジンからの結果文字列
     * @param {boolean} alertOnEmpty - 空結果（ドキュメントなし）でアラートを出すか
     * @returns {void}
     */
    function applyArtboardResult(paletteUI, result, alertOnEmpty) {
        var rulerUnit = getUnitInfo();
        paletteUI.widthUnitText.text = rulerUnit.label;
        paletteUI.heightUnitText.text = rulerUnit.label;

        var resultFields = result ? result.split(ARTBOARD_FIELD_SEPARATOR) : [];
        if (resultFields.length < 4) {
            paletteUI.artboardInfoText.text = "—";
            paletteUI.widthInput.text = "";
            paletteUI.heightInput.text = "";
            if (alertOnEmpty) alert(getLabel("alert.noDocument"));
            return;
        }
        var infoSeparator = (uiLang === "ja") ? "：" : ": ";
        paletteUI.artboardInfoText.text = "#" + resultFields[0] + infoSeparator + resultFields[1];
        paletteUI.widthInput.text = pointToUnitText(parseFloat(resultFields[2]), rulerUnit);
        paletteUI.heightInput.text = pointToUnitText(parseFloat(resultFields[3]), rulerUnit);
    }

    /**
     * アートボード操作をメインエンジンへ委譲し、結果を UI へ反映する
     * @param {object} paletteUI - パレットのコントロール一式
     * @param {string} operation - "read" | "round" | "resize"
     * @param {number} widthPoint - resize 時の幅（pt）
     * @param {number} heightPoint - resize 時の高さ（pt）
     * @param {boolean} alertOnEmpty - 空結果でアラートを出すか
     * @returns {void}
     */
    function runArtboardOperation(paletteUI, operation, widthPoint, heightPoint, alertOnEmpty) {
        var anchorIndex = paletteUI.anchorWidget.selectedAnchorIndex;
        delegateToMainEngine(buildArtboardBridgeCode(operation, widthPoint, heightPoint, anchorIndex), function (result) {
            applyArtboardResult(paletteUI, result, alertOnEmpty);
        });
    }

    /**
     * 現在のアートボード情報を取得して UI へ反映する
     * @param {object} paletteUI - パレットのコントロール一式
     * @returns {void}
     */
    function refreshArtboardInfo(paletteUI) {
        runArtboardOperation(paletteUI, "read", 0, 0, false);
    }

    /**
     * 幅・高さの入力値でアクティブアートボードをリサイズする（不正値は現在値へ戻す）
     * @param {object} paletteUI - パレットのコントロール一式
     * @returns {void}
     */
    function resizeArtboardFromFields(paletteUI) {
        var rulerUnit = getUnitInfo();
        var widthValue = parseFloat(paletteUI.widthInput.text);
        var heightValue = parseFloat(paletteUI.heightInput.text);
        if (isNaN(widthValue) || isNaN(heightValue) || widthValue <= 0 || heightValue <= 0) {
            refreshArtboardInfo(paletteUI);
            return;
        }
        runArtboardOperation(paletteUI, "resize", widthValue * rulerUnit.pointsPerUnit, heightValue * rulerUnit.pointsPerUnit, true);
    }

    // =========================================
    // イベント設定 / Event wiring
    // =========================================

    /**
     * 表示設定（プリセット・アートボード名・枠線・オプション）のイベントを設定する
     * @param {object} paletteUI - パレットのコントロール一式
     * @returns {void}
     */
    function wireDisplaySettingEvents(paletteUI) {
        var i;
        var applyDisplaySettings = function () { applyArtboardDisplaySettings(paletteUI); };

        /* プリセット：UI へ展開してから環境設定へ反映 / Presets: expand into the UI, then write */
        for (i = 0; i < PRESET_KEYS.length; i++) {
            paletteUI.presetRadios[i].presetKey = PRESET_KEYS[i];
            paletteUI.presetRadios[i].onClick = function () {
                applyPresetToUI(paletteUI, this.presetKey);
                applyArtboardDisplaySettings(paletteUI);
                applyOptionSettings(paletteUI);
            };
        }

        paletteUI.showNameCheckbox.onClick = applyDisplaySettings;
        paletteUI.borderColorList.onChange = applyDisplaySettings;
        for (i = 0; i < paletteUI.borderWidthRadios.length; i++) {
            paletteUI.borderWidthRadios[i].onClick = applyDisplaySettings;
        }

        // paletteUI.printBleedCheckbox.onClick = function () { applyOptionSettings(paletteUI); }; /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
        paletteUI.moveLockedHiddenCheckbox.onClick = function () { applyOptionSettings(paletteUI); };
    }

    /**
     * アートボード操作・下部ボタン・ウィンドウのイベントを設定する
     * @param {object} paletteUI - パレットのコントロール一式
     * @returns {void}
     */
    function wireArtboardAndWindowEvents(paletteUI) {
        /* ピクセルグリッドに最適化：XYWH を整数値へ丸める / Optimize: round XYWH to integers */
        paletteUI.optimizePixelGridButton.onClick = function () { runArtboardOperation(paletteUI, "round", 0, 0, true); };

        /* 幅・高さの確定でアートボードをリサイズ / Resize the artboard when width/height are committed */
        paletteUI.widthInput.onChange = function () { resizeArtboardFromFields(paletteUI); };
        paletteUI.heightInput.onChange = function () { resizeArtboardFromFields(paletteUI); };

        /* カンバスカラーの変更（uiCanvasIsWhite: 1=白 / 0=グレー）/ Toggle the canvas color */
        paletteUI.canvasColorButton.onClick = function () {
            appPreferences.setIntegerPreference("uiCanvasIsWhite", readPref("Integer", "uiCanvasIsWhite", 0) === 1 ? 0 : 1);
            refreshArtboardDisplay();
        };
        paletteUI.videoRulerButton.onClick = function () { runMenuCommand("videoruler"); };

        /* アクティブ時に Esc で閉じる / Close on Esc while active */
        paletteUI.paletteWindow.addEventListener("keydown", function (event) {
            if (event.keyName === "Escape") paletteUI.paletteWindow.close();
        });

        /* 再アクティブ時：外部変更とアートボードの切り替えに追従 / On re-activate: follow external changes */
        paletteUI.paletteWindow.onActivate = function () {
            reflectPreferences(paletteUI);
            refreshArtboardInfo(paletteUI);
        };

        /* 常駐エンジンの参照は閉じたらクリア / Clear the persistent-engine reference on close */
        paletteUI.paletteWindow.onClose = function () {
            $.global[PALETTE_GLOBAL_KEY] = null;
        };
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を取得する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義（カテゴリ分け）/ Japanese-English label definitions (categorized) */
    var LABELS = {
        dialog: {
            title: { ja: "アートボード関連の環境設定", en: "Artboard-Related Preferences" }
        },
        panel: {
            currentArtboard: { ja: "現在のアートボード", en: "Current Artboard" },
            artboardDisplay: { ja: "アートボード名と枠線", en: "Artboard Name & Border" },
            artboardBorder: { ja: "アートボードの枠線", en: "Artboard Border" },
            options: { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            borderColor: { ja: "ハイライトのカラー", en: "Highlight Color" },
            borderWidth: { ja: "ストロークの幅", en: "Stroke Width" }
        },
        checkbox: {
            showArtboardName: { ja: "アートボード名を表示", en: "Show Artboard Name" },
            /* showPrintBleedAI は現在未使用（PRINT_BLEED_WIDGET を参照）/ showPrintBleedAI is currently unused (see the PRINT_BLEED_WIDGET note) */
            showPrintBleedAI: {
                ja: "「裁ち落としを印刷」生成AIボタンを表示",
                en: "Show the \"Print Bleed\" Generative AI Button"
            },
            moveLockedHidden: {
                ja: "ロックまたは非表示オブジェクトを一緒に移動",
                en: "Move Locked or Hidden Objects Together"
            }
        },
        radio: {
            preset: {
                /* "default" は ES3 予約語のため引用符付きキーにする / "default" is an ES3 reserved word, so quote the key */
                "default": { ja: "デフォルト", en: "Default" },
                emphasis: { ja: "強調", en: "Emphasis" },
                light: { ja: "ライト", en: "Light" }
            }
        },
        dropdown: {
            borderColor: {
                lightBlue: { ja: "ライトブルー", en: "Light Blue" },
                lightRed: { ja: "サーモンピンク", en: "Light Red" },
                green: { ja: "グリーン", en: "Green" },
                mediumBlue: { ja: "ミディアムブルー", en: "Medium Blue" },
                magenta: { ja: "マゼンタ", en: "Magenta" },
                cyan: { ja: "シアン", en: "Cyan" },
                lightGray: { ja: "ライトグレー", en: "Light Gray" },
                black: { ja: "ブラック", en: "Black" },
                yellow: { ja: "イエロー", en: "Yellow" }
            }
        },
        button: {
            optimizePixelGrid: { ja: "ピクセルグリッドに最適化", en: "Optimize to Pixel Grid" },
            canvasColor: { ja: "カンバスカラーの変更", en: "Change Canvas Color" },
            videoRuler: { ja: "ビデオ定規", en: "Video Ruler" }
        },
        tooltip: {
            sizeField: {
                ja: "値を確定すると、アクティブなアートボードを右の基準点を基準にリサイズします。↑↓で±1、shift併用で±10、option併用で±0.1。",
                en: "Commit a value to resize the active artboard around the reference point on the right. Arrow keys: ±1, Shift ±10, Option ±0.1."
            },
            anchor: { ja: "リサイズの基準点です。", en: "Reference point for resizing." },
            showArtboardName: { ja: "カンバス上にアートボード名を表示します。", en: "Shows the artboard names on the canvas." },
            borderColor: { ja: "アートボードの境界線の色です。", en: "Color of the artboard borders." },
            borderWidth: { ja: "アートボードの境界線の太さです。", en: "Width of the artboard borders." },
            preset: { ja: "まとめて切り替える表示設定の組み合わせです。", en: "A set of display settings applied together." },
            moveLockedHidden: {
                ja: "ロックや非表示のオブジェクトも、アートボードと一緒に動かします。",
                en: "Moves locked and hidden objects along with the artboard."
            },
            optimizePixelGrid: {
                ja: "アクティブなアートボードの位置とサイズを整数値に丸めます。",
                en: "Rounds the active artboard's position and size to whole numbers."
            },
            canvasColor: {
                ja: "アートボード外のカンバスを、白とグレーで切り替えます。",
                en: "Toggles the canvas outside the artboards between white and gray."
            },
            videoRuler: { ja: "ビデオ定規の表示／非表示を切り替えます。", en: "Shows or hides the video ruler." }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    /**
     * ドット区切りキーで LABELS を辿り、現在言語の文言を返す（{slash}→/）
     * @param {string} labelPath - "panel.options" のようなドット区切りキー
     * @returns {string} 現在言語の文言（見つからない場合はキーをそのまま返す）
     */
    function getLabel(labelPath) {
        var pathParts = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathParts.length; i++) {
            if (labelNode == null) return labelPath;
            labelNode = labelNode[pathParts[i]];
        }
        if (labelNode == null) return labelPath;
        var text = labelNode[uiLang] || labelNode.en || "";
        return text.replace(/\{slash\}/g, "/");
    }

    /**
     * コロン付きの項目名を返す（日本語は全角「：」、英語は半角「:」）
     * @param {string} labelPath - LABELS のドット区切りキー
     * @returns {string} コロンを付けた文言
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // メイン処理 / Main process
    // =========================================

    /**
     * すでに開いているパレットを返す（無効な参照・表示前の残骸は null）
     * 表示中のものだけを有効と見なす。構築途中でエラーになったウィンドウを掴むと、
     * イベント未配線のパレットを開き続けることになるため
     * 閉じて破棄されたウィンドウは visible の参照で例外になるため try で受ける
     * @returns {Window|null} 既存のパレット
     */
    function getExistingPalette() {
        try {
            var palette = $.global[PALETTE_GLOBAL_KEY];
            return (palette && palette.visible === true) ? palette : null;
        } catch (e) {
            return null;
        }
    }

    /**
     * パレットを構築して表示する（重複起動時は既存のパレットを前面に出す）
     * @returns {void}
     */
    function main() {
        var existingPalette = getExistingPalette();
        if (existingPalette) {
            existingPalette.show();
            return;
        }

        var paletteUI = buildPalette();
        wireDisplaySettingEvents(paletteUI);
        wireArtboardAndWindowEvents(paletteUI);
        reflectPreferences(paletteUI);
        refreshArtboardInfo(paletteUI);

        paletteUI.paletteWindow.center();
        paletteUI.paletteWindow.show();

        /* 構築と表示がすべて通ってから参照を保持する（途中で失敗した窓を残さない）
           Store the reference only after everything succeeded, so a half-built window is never kept */
        $.global[PALETTE_GLOBAL_KEY] = paletteUI.paletteWindow;

        /* レイアウト確定後にボタン高さを 2px 詰める / Trim the button heights by 2px after layout */
        trimButtonHeight(paletteUI.optimizePixelGridButton, 2);
    }

    main();

}());
