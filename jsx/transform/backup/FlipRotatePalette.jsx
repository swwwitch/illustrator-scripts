#target illustrator
#targetengine "FlipRotatePalette"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトの反転・回転を、9軸の基準点を指定しながら、常駐パレットのアイコンからまとめて操作するユーティリティです。操作した時点で即時反映されます。

- 左右反転／上下反転／回転（反時計回り・時計回り）をアイコンボタンで実行
- 9軸（3×3）の基準点ウィジェットで反転・回転の基点を指定（既定は中央）。基点は選択全体の可視バウンディングを基準に算出
- 回転角は 45°（ROTATE_ANGLE で変更可）
- DOM 変形は BridgeTalk でメインエンジンへ委譲し、適用後すぐに再描画（常駐パレットは DOM 接続を失うため）
- ライト／ダーク UI に合わせてアイコンの配色を自動切り替え（ライトは枠線＋ホバーで背景）

### Overview

A persistent-palette utility that flips and rotates the selection from icon buttons, with a selectable 9-axis pivot. Every action applies immediately.

- Flip horizontal / vertical and rotate (counterclockwise / clockwise) from icon buttons
- A 9-axis (3x3) anchor widget sets the pivot for flips and rotations (center by default); the pivot is computed from the selection's overall visible bounds
- Rotation angle is 45 degrees (configurable via ROTATE_ANGLE)
- DOM transforms are delegated to the main engine via BridgeTalk and redrawn immediately (a persistent palette loses its DOM connection)
- Icon colors adapt to the light / dark UI (light shows a border, with a hover background)

### 更新履歴 / Change Log

- v1.0.0 (20260630): 初期バージョン。反転・回転アイコン、9軸の基準点、ライト／ダーク対応、BridgeTalk でメインエンジンへ委譲。

*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.0.0";

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var ROTATE_ANGLE = 45;          /* 回転アイコンの回転角（度）。正＝反時計回り / Rotate icon angle (deg); positive = CCW */
    var COLUMNS_PER_ROW = 2;        /* 何個ごとに改行するか / Icons per row before wrapping */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 現在のUI言語を取得 / Get the current UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* 日英ラベル定義（カテゴリ分け）/ Japanese-English label definitions (categorized) */
    var LABELS = {
        dialog: {
            title: { ja: "反転・回転", en: "Flip & Rotate" }
        },
        panel: {
            transform: { ja: "反転・回転", en: "Flip & Rotate" }
        },
        tooltip: {
            flipHorizontal: { ja: "左右反転", en: "Flip horizontal" },
            flipVertical:   { ja: "上下反転", en: "Flip vertical" },
            rotateCCW:      { ja: "回転（反時計回り）", en: "Rotate counterclockwise" },
            rotateCW:       { ja: "回転（時計回り）", en: "Rotate clockwise" },
            anchor:         { ja: "基準点（反転・回転の基点）", en: "Anchor point (pivot for flip / rotate)" }
        }
    };

    /* ドット区切りキーで LABELS を辿り現在言語の文言を返す（{slash}→/）/ Resolve a dot-path key in LABELS to the current-language text ({slash}→/) */
    function getLabel(key) {
        var parts = key.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (node == null) return key;
            node = node[parts[i]];
        }
        if (node == null) return key;
        var text = node[currentLanguage] || node.en || "";
        return text.replace(/\{slash\}/g, "/");
    }

    /* 現在言語のラベル文字列を返す / Return the current-language label string */
    function L(key) {
        return getLabel(key);
    }

    // =========================================
    // 状態 / State
    // =========================================

    /* 基準点（9軸）: 0..8 を行優先（0=左上, 4=中央, 8=右下）。反転・回転の基点に使う / 9-axis anchor 0..8 row-major (0=top-left, 4=center, 8=bottom-right); used as the flip/rotate pivot */
    var selectedAnchorIndex = 4;

    var iconButtonDefs = [
        { name: "FLIP_HORIZONTAL", icon: "flipHorizontal", tooltip: "tooltip.flipHorizontal" },
        { name: "FLIP_VERTICAL",   icon: "flipVertical",   tooltip: "tooltip.flipVertical" },
        { name: "ROTATE",          icon: "rotate",         tooltip: "tooltip.rotateCCW" },
        { name: "ROTATE_FLIP",     icon: "rotateFlip",     tooltip: "tooltip.rotateCW" }
    ];

    // =========================================
    // パレット構築 / Build the palette
    // =========================================

    /* すでにパレットが開いていれば前面に出して終了 / If the palette already exists, bring it forward and return */
    try {
        if ($.global.flipRotatePaletteWindow && $.global.flipRotatePaletteWindow.visible) {
            $.global.flipRotatePaletteWindow.show();
            return;
        }
    } catch (e) {
    }

    /* UI の明暗に合わせてアイコン色・背景色・枠線を決める / Decide icon, background and border colors from the light/dark UI */
    var lightUI = isLightUI();
    var uiBrightness = getUIBrightness();
    var iconColor       = lightUI ? [0.25, 0.25, 0.25, 1] : [0.85, 0.85, 0.85, 1];
    /* ライトモードは背景を塗らず（ウィンドウ色）薄いグレーの枠を付ける。ダークは枠なし / Light: no fill (window color) + light gray border; dark: no border */
    var iconBorderColor = lightUI ? [0.65, 0.65, 0.65, 1] : null;
    var iconBaseBg      = lightUI ? grayColor(uiBrightness)        : [0.28, 0.28, 0.28, 1];
    /* マウスオーバー時の背景（ライトは少し暗く、ダークは少し明るく）/ Hover background (slightly darker in light, lighter in dark) */
    var iconHoverBg     = lightUI ? grayColor(uiBrightness - 0.10) : [0.38, 0.38, 0.38, 1];

    var paletteWindow = new Window("palette", L('dialog.title') + ' ' + SCRIPT_VERSION, undefined, { resizeable: false });
    $.global.flipRotatePaletteWindow = paletteWindow;

    paletteWindow.orientation = "column";
    paletteWindow.alignChildren = ["center", "top"];
    paletteWindow.margins = 12;   // 全体の外側余白
    paletteWindow.spacing = 0;

    /* 「反転・回転」パネル（ボタン類をまとめる枠）/ "Flip & Rotate" panel that frames the buttons */
    var mainPanel = paletteWindow.add("panel", undefined, L('panel.transform'));
    mainPanel.orientation = "column";
    mainPanel.alignChildren = ["center", "top"];
    mainPanel.margins = 10;
    mainPanel.spacing = 0;

    var buttonGrid = mainPanel.add("group");
    buttonGrid.orientation = "column";
    buttonGrid.alignChildren = ["center", "center"];
    buttonGrid.spacing = 7;

    var currentRow = null;
    for (var i = 0; i < iconButtonDefs.length; i++) {
        if ((i % COLUMNS_PER_ROW) === 0) {
            currentRow = buttonGrid.add("group");
            currentRow.orientation = "row";
            currentRow.alignChildren = ["center", "center"];
            currentRow.spacing = 7;
        }
        addIconButton(currentRow, iconButtonDefs[i]);
    }

    // 基準点（9軸）ウィジェット
    var anchorRow = mainPanel.add("group");
    anchorRow.orientation = "row";
    anchorRow.alignChildren = ["center", "center"];
    anchorRow.margins = [0, 6, 0, 0];
    addAnchorWidget(anchorRow);

    paletteWindow.onClose = function () {
        $.global.flipRotatePaletteWindow = null;
    };

    paletteWindow.show();

    /* アイコンボタンを1つ生成して登録 / Create and register one icon button */
    function addIconButton(parentGroup, buttonDef) {
        var button = parentGroup.add("button", undefined, "");
        button.helpTip = L(buttonDef.tooltip);
        button.preferredSize = [26, 26];
        button.minimumSize = [26, 26];
        button.maximumSize = [26, 26];
        button.iconType = buttonDef.icon;
        button.isHover = false;
        button.onDraw = function () {
            drawIcon(this);
        };
        button.onClick = function () {
            handleIconAction(buttonDef.name);
        };
        // マウスオーバーで背景を付けるためホバー状態を切り替える / Toggle hover to show a background on mouseover
        try {
            button.addEventListener("mouseover", function () { button.isHover = true; redrawControl(button); });
            button.addEventListener("mouseout", function () { button.isHover = false; redrawControl(button); });
        } catch (e) {
        }
    }

    /* コントロールを再描画（notify は環境により例外を投げ得るので保護）/ Redraw a control (notify can throw in some environments) */
    function redrawControl(control) {
        try { control.notify("onDraw"); } catch (e) {}
    }

    /* アイコンに対応する変形を実行する / Run the transform for an icon */
    function handleIconAction(name) {
        if (name === "FLIP_HORIZONTAL") {
            applyFlip(-100, 100);
        } else if (name === "FLIP_VERTICAL") {
            applyFlip(100, -100);
        } else if (name === "ROTATE") {
            applyRotate(ROTATE_ANGLE);     // 反時計回り
        } else if (name === "ROTATE_FLIP") {
            applyRotate(-ROTATE_ANGLE);    // 時計回り
        }
    }

    /* 反転（メインエンジンへ委譲）/ Flip (delegated to the main engine) */
    function applyFlip(scaleX, scaleY) {
        btFlipSelection(scaleX, scaleY);
    }

    /* 回転（メインエンジンへ委譲）/ Rotate (delegated to the main engine) */
    function applyRotate(angleDegrees) {
        btRotateSelection(angleDegrees);
    }

    /* 9軸（3×3）の基準点ウィジェットを生成 / Create the 9-axis (3x3) anchor widget */
    function addAnchorWidget(parentGroup) {
        var widget = parentGroup.add("button", undefined, "");
        widget.helpTip = L('tooltip.anchor');
        widget.preferredSize = [44, 44];
        widget.minimumSize = [44, 44];
        widget.maximumSize = [44, 44];
        widget.onDraw = function () {
            drawAnchorWidget(this);
        };
        widget.onClick = function () {
            // クリック位置の判定は mousedown 側で行う（onClick は何もしない）。
        };
        // クリックした 3×3 のセルを基準点に設定（クリック座標はコントロール基準）。
        try {
            widget.addEventListener("mousedown", function (event) {
                var cellWidth = widget.size[0] / 3;
                var cellHeight = widget.size[1] / 3;
                var col = Math.floor(event.clientX / cellWidth);
                var row = Math.floor(event.clientY / cellHeight);
                if (col < 0) col = 0;
                if (col > 2) col = 2;
                if (row < 0) row = 0;
                if (row > 2) row = 2;
                selectedAnchorIndex = row * 3 + col;
                redrawControl(widget);
            });
        } catch (e) {
        }
        return widget;
    }

    /* 9軸ウィジェットを描画（外周の□をケイ線でつなぐ・中央は独立）/ Draw the 9-axis widget (outer squares joined by rules; center stands alone) */
    function drawAnchorWidget(widget) {
        var graphics = widget.graphics;
        var width = widget.size[0];
        var height = widget.size[1];

        try {
            // 9軸ウィジェットは枠線なし（背景のみ、ウィンドウ色に馴染ませる）。
            graphics.rectPath(0, 0, width, height);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, iconBaseBg));
        } catch (e0) {
        }

        var cellSize = 6;   // 四角のサイズ
        var cellGap = 5;    // 四角どうしの間隔
        var cellStep = cellSize + cellGap;
        var gridSize = cellSize * 3 + cellGap * 2;
        var originX = Math.round((width - gridSize) / 2);
        var originY = Math.round((height - gridSize) / 2);

        function anchorCellX(index) { return originX + (index % 3) * cellStep; }
        function anchorCellY(index) { return originY + Math.floor(index / 3) * cellStep; }

        // 中央(4)を除く外周の□どうしをケイ線（隙間部分）でつなぐ。
        var connections = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];
        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, iconColor, 1);
        for (var i = 0; i < connections.length; i++) {
            var cellA = connections[i][0];
            var cellB = connections[i][1];
            var cellAX = anchorCellX(cellA);
            var cellAY = anchorCellY(cellA);
            var cellBX = anchorCellX(cellB);
            var cellBY = anchorCellY(cellB);
            graphics.newPath();
            if (cellB - cellA === 1) {
                // 横方向：右隣の□へ / Horizontal: to the square on the right
                graphics.moveTo(cellAX + cellSize, cellAY + cellSize / 2);
                graphics.lineTo(cellBX, cellBY + cellSize / 2);
            } else {
                // 縦方向：下の□へ / Vertical: to the square below
                graphics.moveTo(cellAX + cellSize / 2, cellAY + cellSize);
                graphics.lineTo(cellBX + cellSize / 2, cellBY);
            }
            graphics.strokePath(linePen);
        }

        for (var index = 0; index < 9; index++) {
            drawAnchorCell(graphics, anchorCellX(index), anchorCellY(index), cellSize, index === selectedAnchorIndex);
        }
    }

    /* 基準点セルの□を1つ描画（選択中は塗り、非選択は枠線）/ Draw one anchor-cell square (filled if selected, outlined otherwise) */
    function drawAnchorCell(graphics, x, y, size, selected) {
        graphics.newPath();
        graphics.moveTo(x, y);
        graphics.lineTo(x + size, y);
        graphics.lineTo(x + size, y + size);
        graphics.lineTo(x, y + size);
        graphics.closePath();
        if (selected) {
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, iconColor));
        } else {
            graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, iconColor, 1));
        }
    }

    /* アイコンボタンの背景・枠線を描き、種類に応じた図柄を描画 / Draw the icon button background/border and dispatch to the right glyph */
    function drawIcon(button) {
        var graphics = button.graphics;
        var width = button.size[0];
        var height = button.size[1];
        var backgroundColor = (button.isHover === true) ? iconHoverBg : iconBaseBg;

        try {
            graphics.rectPath(0, 0, width, height);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));

            // ライトモードは薄いグレーの枠線を描く。
            if (iconBorderColor) {
                graphics.rectPath(0.5, 0.5, width - 1, height - 1);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, iconBorderColor, 1));
            }
        } catch (e1) {
            try {
                graphics.drawOSControl();
            } catch (e2) {
            }
        }

        if (button.iconType === "rotate") {
            drawRotateIcon(graphics, width, height, false);
        } else if (button.iconType === "rotateFlip") {
            drawRotateIcon(graphics, width, height, true);
        } else {
            drawFlipIcon(graphics, button.iconType, width, height);
        }
    }

    /* 反転アイコン（軸の点線＋向かい合う三角形）を描画 / Draw a flip icon (dotted axis + opposing triangles) */
    function drawFlipIcon(graphics, iconType, width, height) {
        var color = iconColor;
        var centerX = width / 2;
        var centerY = height / 2;

        if (iconType === "flipVertical") {
            // 横の点線を軸に、上向き／下向きの三角形を配置する。
            drawDottedLine(graphics, 5, centerY, width - 5, centerY, color);
            drawTriangle(graphics, [[centerX - 5, 4], [centerX + 5, 4], [centerX, centerY - 2]], color, true);
            drawTriangle(graphics, [[centerX - 5, height - 4], [centerX + 5, height - 4], [centerX, centerY + 2]], color, false);
        } else {
            // 縦の点線を軸に、左向き／右向きの三角形を配置する。
            drawDottedLine(graphics, centerX, 5, centerX, height - 5, color);
            drawTriangle(graphics, [[4, centerY - 5], [4, centerY + 5], [centerX - 2, centerY]], color, true);
            drawTriangle(graphics, [[width - 4, centerY - 5], [width - 4, centerY + 5], [centerX + 2, centerY]], color, false);
        }
    }

    /* 回転アイコン（一定幅の実線弧＋点線弧＋矢じり）を描画。mirror で左右反転 / Draw a rotate icon (uniform-width solid arc + dotted arc + arrowhead); mirror flips it horizontally */
    function drawRotateIcon(graphics, width, height, mirror) {
        var color = iconColor;
        var centerX = width / 2;
        var centerY = height / 2 + 1;
        var radius = 7.5;
        var mirrorSign = mirror ? -1 : 1;   // 左右反転のときは x をミラーする / Mirror x when flipped
        var headDeg = 232;                  // 矢じりの位置（左上）/ Arrowhead position (top-left)

        // 実線の弧は一定幅・単一パスでなめらかに描く（矢じりで頭の重さを表現）。
        strokeArc(graphics, color, centerX, centerY, radius, headDeg, 410, mirrorSign, 1.8, 1.8);
        // 下側は四角い点線。
        drawDottedArc(graphics, color, centerX, centerY, radius, 50, 150, mirrorSign, 1.8);

        // 左上（反転時は右上）に大きめの矢じりを付ける。矢じりの先（左下）は空白。
        var headRad = headDeg * Math.PI / 180;
        var headX = centerX + radius * Math.cos(headRad);
        var headY = centerY + radius * Math.sin(headRad);
        var tangentX = Math.sin(headRad);   // 反時計回り（角度が減る向き）の接線
        var tangentY = -Math.cos(headRad);
        var perpX = -tangentY;
        var perpY = tangentX;
        var tipForward = 4;       // 矢じり先端の前方への張り出し / Arrowhead tip extent (forward)
        var tipBack = 2;          // 矢じり後方への張り出し / Arrowhead extent (backward)
        var tipHalfWidth = 4.5;   // 矢じりの片側の幅 / Arrowhead half width

        var arrowPoints = [
            [headX + tangentX * tipForward, headY + tangentY * tipForward],
            [headX - tangentX * tipBack + perpX * tipHalfWidth, headY - tangentY * tipBack + perpY * tipHalfWidth],
            [headX - tangentX * tipBack - perpX * tipHalfWidth, headY - tangentY * tipBack - perpY * tipHalfWidth]
        ];

        if (mirror) {
            for (var i = 0; i < arrowPoints.length; i++) {
                arrowPoints[i][0] = 2 * centerX - arrowPoints[i][0];
            }
        }

        drawTriangle(graphics, arrowPoints, color, true);
    }

    /* 円弧を線分で近似して描く。線幅一定なら継ぎ目の出ない単一パス、変化させる場合のみセグメントごとに描く（先細り）/
       Stroke an arc as line segments: a single seamless path when the width is constant, per-segment only when it tapers */
    function strokeArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign, startWidth, endWidth) {
        var segments = Math.max(8, Math.round(Math.abs(endDeg - startDeg) / 5));

        function arcX(ratio) { return centerX + mirrorSign * radius * Math.cos((startDeg + (endDeg - startDeg) * ratio) * Math.PI / 180); }
        function arcY(ratio) { return centerY + radius * Math.sin((startDeg + (endDeg - startDeg) * ratio) * Math.PI / 180); }

        /* 幅一定：1本の連続パスで描くのでガタつかない / Constant width: one continuous path, so no jaggedness */
        if (startWidth === endWidth) {
            graphics.newPath();
            graphics.moveTo(arcX(0), arcY(0));
            for (var i = 1; i <= segments; i++) {
                graphics.lineTo(arcX(i / segments), arcY(i / segments));
            }
            graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, startWidth));
            return;
        }

        /* 幅可変：セグメントごとに線幅を補間 / Variable width: interpolate per segment */
        for (var j = 1; j <= segments; j++) {
            var ratio = j / segments;
            var lineWidth = startWidth + (endWidth - startWidth) * ratio;
            graphics.newPath();
            graphics.moveTo(arcX((j - 1) / segments), arcY((j - 1) / segments));
            graphics.lineTo(arcX(ratio), arcY(ratio));
            graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, lineWidth));
        }
    }

    /* 円弧に沿って四角い点線を描く / Draw a square-dotted line along an arc */
    function drawDottedArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign, dotWidth) {
        var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, dotWidth);
        var stepDeg = 13;
        var dashHalf = 0.9;

        for (var deg = startDeg; deg <= endDeg; deg += stepDeg) {
            var rad = deg * Math.PI / 180;
            var x = centerX + mirrorSign * radius * Math.cos(rad);
            var y = centerY + radius * Math.sin(rad);
            var tangentX = mirrorSign * Math.sin(rad);
            var tangentY = -Math.cos(rad);

            graphics.newPath();
            graphics.moveTo(x - tangentX * dashHalf, y - tangentY * dashHalf);
            graphics.lineTo(x + tangentX * dashHalf, y + tangentY * dashHalf);
            graphics.strokePath(pen);
        }
    }

    /* 3点の三角形を塗り or 線で描く / Draw a triangle from 3 points, filled or stroked */
    function drawTriangle(graphics, points, color, fill) {
        graphics.newPath();
        graphics.moveTo(points[0][0], points[0][1]);
        graphics.lineTo(points[1][0], points[1][1]);
        graphics.lineTo(points[2][0], points[2][1]);
        graphics.closePath();
        if (fill) {
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
        } else {
            graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, 1.2));
        }
    }

    /* 水平または垂直の点線を描く / Draw a horizontal or vertical dotted line */
    function drawDottedLine(graphics, x1, y1, x2, y2, color) {
        var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, 1);
        var isHorizontal = (y1 === y2);
        var totalLength = isHorizontal ? (x2 - x1) : (y2 - y1);
        var dashStep = 3;

        for (var pos = 0; pos < totalLength; pos += dashStep) {
            graphics.newPath();
            if (isHorizontal) {
                graphics.moveTo(x1 + pos, y1);
                graphics.lineTo(Math.min(x1 + pos + 1.5, x2), y1);
            } else {
                graphics.moveTo(x1, y1 + pos);
                graphics.lineTo(x1, Math.min(y1 + pos + 1.5, y2));
            }
            graphics.strokePath(pen);
        }
    }

    // =========================================
    // UI の明暗判定 / Light vs. dark UI
    // =========================================

    /* UI 明度(0..1)を取得、取得失敗時は 1（明るい）を返す / Get UI brightness (0..1); 1 on failure */
    function getUIBrightness() {
        try {
            var brightness = app.preferences.getRealPreference("uiBrightness");
            if (brightness < 0) brightness = 0;
            if (brightness > 1) brightness = 1;
            return brightness;
        } catch (e) {
            return 1;
        }
    }

    /* UI 明度 > 0.5 ならライト、取得失敗時はダーク扱い / Light if uiBrightness > 0.5; dark on failure */
    function isLightUI() {
        try {
            return app.preferences.getRealPreference("uiBrightness") > 0.5;
        } catch (e) {
            return false;
        }
    }

    /* グレーの RGBA を返す（0..1 にクランプ）/ Build a clamped gray RGBA */
    function grayColor(value) {
        if (value < 0) value = 0;
        if (value > 1) value = 1;
        return [value, value, value, 1];
    }

    // =========================================
    // BridgeTalk 委譲（DOM 操作）/ BridgeTalk delegation (DOM operations)
    // 常駐パレットエンジンは DOM 接続を失うため、変形はメインエンジンへ委譲する
    // A persistent palette engine loses its DOM connection, so transforms are delegated to the main engine
    // =========================================

    /* メインエンジン（target="illustrator"）へコードを送って実行 / Send code to the main engine and run it */
    function runInMainEngine(bodyCode) {
        try {
            var bridge = new BridgeTalk();
            bridge.target = "illustrator"; /* #targetengine 指定なし＝メインエンジン / no engine = main engine */
            bridge.body = bodyCode;
            bridge.onError = function (message) {
                try { $.writeln("FlipRotatePalette BridgeTalk error: " + message.body); } catch (e) {}
            };
            bridge.send();
        } catch (e) {
            /* BridgeTalk 不可時は同一エンジンで直接実行 / Fallback: run directly in this engine */
            try {
                eval(bodyCode);
            } catch (e2) {
            }
        }
    }

    /* 選択した 9 軸の基準点に対応する anchorX/anchorY の式（body 内の left/right/top/bottom を参照）を返す */
    /* Return the anchorX/anchorY expressions for the selected 9-axis point (referencing left/right/top/bottom in the body) */
    function getAnchorExpressions() {
        var col = selectedAnchorIndex % 3;
        var row = Math.floor(selectedAnchorIndex / 3);
        var xExpr = (col === 0) ? "left" : (col === 1) ? "((left+right)/2)" : "right";
        var yExpr = (row === 0) ? "top" : (row === 1) ? "((top+bottom)/2)" : "bottom";
        return "var anchorX=" + xExpr + ",anchorY=" + yExpr + ";";
    }

    /* 共通の委譲本体：バウンディング測定→基準点→行列適用。matrixCode で matrix を組み立てる
       Shared delegation body: measure bounds -> anchor -> apply matrix; matrixCode builds the matrix */
    function btTransformSelection(matrixCode) {
        var body = '' +
            'if(app.documents.length>0){' +
            'var doc=app.activeDocument,selection=doc.selection;' +
            'if(selection&&selection.length>0){' +
            'var left=Infinity,top=-Infinity,right=-Infinity,bottom=Infinity,measured=false;' +
            'for(var i=0;i<selection.length;i++){try{var b=selection[i].visibleBounds;if(b[0]<left)left=b[0];if(b[1]>top)top=b[1];if(b[2]>right)right=b[2];if(b[3]<bottom)bottom=b[3];measured=true;}catch(e){}}' +
            'if(measured){' +
            getAnchorExpressions() +
            'var matrix=app.getIdentityMatrix();' + matrixCode +
            'for(var i=0;i<selection.length;i++){try{selection[i].transform(matrix,true,true,true,true,1,Transformation.DOCUMENTORIGIN);}catch(e){}}' +
            'app.redraw();' +
            '}}}';
        runInMainEngine(body);
    }

    /* 選択を、選択した 9 軸基準点を基点に反転（水平=-100,100／垂直=100,-100）/ Flip about the selected 9-axis anchor */
    /* 係数は数値化して埋め込む（'1-' + (-1) だと "1--1" になりデクリメント解釈で構文エラーになるため）
       Coefficients are precomputed as numbers ('1-' + (-1) would yield "1--1", a decrement → syntax error) */
    function btFlipSelection(scaleX, scaleY) {
        var scaleFractionX = Number(scaleX) / 100; /* -1 or 1 */
        var scaleFractionY = Number(scaleY) / 100;
        btTransformSelection(
            'matrix.mValueA=' + scaleFractionX + ';matrix.mValueD=' + scaleFractionY + ';' +
            'matrix.mValueTX=anchorX*' + (1 - scaleFractionX) + ';matrix.mValueTY=anchorY*' + (1 - scaleFractionY) + ';'
        );
    }

    /* 選択を、選択した 9 軸基準点を基点に回転（angleDegrees：正＝反時計回り／負＝時計回り）/ Rotate about the selected 9-axis anchor (positive = CCW) */
    function btRotateSelection(angleDegrees) {
        var radians = Number(angleDegrees) * Math.PI / 180;
        var cosAngle = Math.cos(radians);
        var sinAngle = Math.sin(radians);
        var oneMinusCos = 1 - cosAngle;   /* 係数は数値化（"1--0.7" のような構文エラーを避ける）/ Precompute to avoid "1--0.7"-style syntax errors */
        btTransformSelection(
            'matrix.mValueA=' + cosAngle + ';matrix.mValueB=' + sinAngle + ';matrix.mValueC=' + (-sinAngle) + ';matrix.mValueD=' + cosAngle + ';' +
            'matrix.mValueTX=anchorX*' + oneMinusCos + '+anchorY*' + sinAngle + ';' +
            'matrix.mValueTY=anchorY*' + oneMinusCos + '-anchorX*' + sinAngle + ';'
        );
    }
}());
