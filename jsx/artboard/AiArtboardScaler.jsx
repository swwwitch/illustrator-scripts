#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボードを現在サイズ基準でスケール変更します。ダイアログを開いたままライブプレビューできます。

- 対象：現在のアートボード／すべて／指定（例: 3, 4 または 3-5、入力は厳密に検証）
- スケール：幅・高さは一意の倍率で相互連動、9つの基準点を固定して再計算
- 幅・高さの入力はアクティブアートボード基準の%に換算される。対象が複数の場合は各アートボードが同率で拡縮（各アートボードを入力した絶対サイズに揃えるものではない）
- ライブプレビューは「アートボードの矩形」のみ（完全可逆）。オブジェクトはOK確定時に一度だけ拡縮（線幅も正しく拡縮）
- オブジェクト連動：複数アートボードにまたがるオブジェクトも所属ルール（参照/uuidで一意化）で1回だけ変形
- オブジェクト連動はロック・非表示のオブジェクトを対象外（そのアートボードの矩形のみ変更されます）
- ピクセルグリッド最適化：基準点固定と幅・高さの整数化を優先（基準点も整数化するため、基準点でない側の端は0.5px単位になる場合あり）。オブジェクトは整数化後の実効倍率で拡縮
- 定規単位（rulerType 0〜10）に追従してサイズ欄を表示
- 初期状態を基準として復元・再適用し、変形誤差の蓄積を抑制。変形・復元失敗時は安全に復元し通知

### 履歴

v1.0.0 2026-07-15 初版 / Initial release

*/

/*

### Overview

Scales artboards based on their current size, with a live preview while the dialog stays open.

- Target: current / all / specify (e.g. 3, 4 or 3-5, strictly validated)
- Scale: uniform ratio; width/height linked, recalculated with a fixed 9-point anchor
- Width/height inputs are converted to a % relative to the active artboard; with multiple targets each artboard scales by the same ratio (it does not set each artboard to the entered absolute size)
- Live preview covers the artboard rectangles only (fully reversible); objects are scaled once at commit (line width scales too)
- Scale objects: objects spanning multiple artboards are assigned once by an ownership rule (deduped by reference/uuid) and transformed only once
- Locked/hidden objects are excluded from object scaling (only the artboard rectangle changes)
- Pixel-grid optimization: prioritizes a fixed anchor and integer width/height (the pivot is integerized too, so non-anchored edges may land on 0.5px); objects use the post-rounding effective ratio
- Size fields follow the current ruler unit (rulerType 0..10)
- Uses the initial state as the baseline for restore/re-apply to limit drift; on transform/restore failure it reverts safely and notifies

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.0";

(function () {

    // =========================================
    // ユーザー設定セクション / User settings
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */

    /* ウィンドウの共通設定 / Apply shared window layout */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 現在のUI言語を返す / Return current UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "アートボードサイズ変更", en: "Resize Artboards" }
        },
        panel: {
            target: { ja: "対象", en: "Target" },
            scale:  { ja: "サイズ・スケール", en: "Size & Scale" }
        },
        radio: {
            current: { ja: "現在のアートボード", en: "Current artboard" },
            all:     { ja: "すべてのアートボード", en: "All artboards" },
            specify: { ja: "指定", en: "Specify" }
        },
        checkbox: {
            scaleObjects: {
                ja: "オブジェクトと一緒に拡大・縮小",
                en: "Scale objects together"
            },
            pixelGrid: {
                ja: "ピクセルグリッドに最適化",
                en: "Optimize for pixel grid"
            }
        },
        field: {
            scale:  { ja: "スケール", en: "Scale" },
            width:  { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" }
        },
        label: {
            anchor: { ja: "基準点", en: "Anchor" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            apply:  { ja: "適用", en: "Apply" }
        },
        alert: {
            noDocument:      { ja: "開いているドキュメントがありません。", en: "No documents are open." },
            invalidNumber:   { ja: "正の数値を入力してください。", en: "Please enter positive numbers." },
            invalidSelection: {
                ja: "対象アートボードの指定が正しくありません（例: 3, 4 または 3-5）。",
                en: "Invalid artboard selection (e.g. 3, 4 or 3-5)."
            },
            transformError: {
                ja: "変形の適用に失敗したため、処理を中止して元に戻しました。",
                en: "Failed to apply the transform; the operation was cancelled and reverted."
            },
            restoreError: {
                ja: "アートボードの復元に失敗しました。手動で元に戻してください（取り消し等）。",
                en: "Failed to restore the artboards. Please revert manually (e.g. Undo)."
            }
        }
    };

    /**
     * ドット区切りパスでラベルを取得しローカライズする（キー欠落時は path を返す）
     * Resolve a localized label by dot path (returns the path itself if a key is missing)
     * @param {string} path 例 "panel.target"
     * @returns {string} ローカライズ済み文字列（見つからなければ path）
     */
    function getLocalizedText(path) {
        var parts = path.split(".");
        var node = LABELS;
        /* 各階層を安全に辿る（途中で欠落したら path を返す） / Walk each level safely; return path if anything is missing */
        for (var i = 0; i < parts.length; i++) {
            if (node === null || typeof node !== "object" || !node.hasOwnProperty(parts[i])) {
                return path;
            }
            node = node[parts[i]];
        }
        /* 現在言語→英語→日本語の順にフォールバック / Fall back current language → English → Japanese */
        var text = node[currentLanguage];
        if (text === undefined || text === null) { text = node.en; }
        if (text === undefined || text === null) { text = node.ja; }
        if (typeof text !== "string") { return path; } /* どの言語値も無ければ path / no language value → path */
        return text.replace(/\{slash\}/g, "/");
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(path) {
        return getLocalizedText(path) + (currentLanguage === "ja" ? "：" : ":");
    }

    // =========================================
    // 単位 / Unit
    // =========================================

    /* rulerType(0..10) → 単位ラベルと1単位あたりのpt数 / rulerType(0..10) → unit label and pt-per-unit */
    var RULER_UNITS = [
        { label: "inch",  factor: 72.0 },              /* 0 */
        { label: "mm",    factor: 72.0 / 25.4 },       /* 1 */
        { label: "pt",    factor: 1.0 },               /* 2 */
        { label: "pica",  factor: 12.0 },              /* 3 */
        { label: "cm",    factor: 72.0 / 2.54 },       /* 4 */
        { label: "Q",     factor: 72.0 / 25.4 * 0.25 },/* 5 */
        { label: "px",    factor: 1.0 },               /* 6 */
        { label: "ft/in", factor: 72.0 * 12.0 },       /* 7  ※内部換算はフィート扱い / treated as feet internally */
        { label: "m",     factor: 72.0 / 0.0254 },     /* 8 */
        { label: "yd",    factor: 72.0 * 36.0 },       /* 9 */
        { label: "ft",    factor: 72.0 * 12.0 }        /* 10 */
    ];

    /**
     * 現在の定規単位（rulerType）からラベルとpt換算係数を取得する / Get the current ruler unit's label and pt factor
     * 未知の値は pt にフォールバックする。 / Unknown values fall back to pt.
     * @returns {object} { label: string, factor: number } factor はその単位1つあたりのpt数
     */
    function getRulerUnit() {
        return RULER_UNITS[app.preferences.getIntegerPreference("rulerType")] || { label: "pt", factor: 1.0 };
    }

    /* 数値を小数2桁に丸めて文字列で返す / Round a number to 2 decimals and return as string */
    function formatNumber(value) {
        return "" + (Math.round(value * 100) / 100);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* ラベル幅を揃えるための固定幅（px。「スケール:」が収まる幅） / Fixed label width to align rows (fits "Scale:") */
    var FIELD_LABEL_WIDTH = 64;

    /**
     * 「ラベル: [入力欄] 単位」の1行を生成する / Build one "label: [input] unit" row
     * @param {Panel} parent 追加先パネル
     * @param {string} labelText ラベル文字列
     * @param {string} defaultValue 入力欄の初期値
     * @param {string} unitLabel 入力欄の後ろに表示する単位
     * @returns {EditText} 生成した入力欄
     */
    function addSizeField(parent, labelText, defaultValue, unitLabel) {
        var row = parent.add("group");
        row.orientation = "row";
        var label = row.add("statictext", undefined, labelText);
        label.preferredSize.width = FIELD_LABEL_WIDTH; /* ラベル幅を固定して揃える / Fix width to align labels */
        label.justify = "right";                       /* 右揃え / Right-align the label */
        var input = row.add("edittext", undefined, defaultValue);
        input.characters = 4;
        row.add("statictext", undefined, unitLabel); /* 入力欄の後ろに単位 / Unit after the input */
        return input;
    }

    /**
     * テキストフィールドで↑↓キーによる値の増減を有効にする / Enable arrow-key value change on a text field
     * ↑↓で±1、Shift併用で±10（10の倍数にスナップ）。optionキーは使わない。
     * @param {EditText} editText 対象の入力欄（数値を保持していること）
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) { return; }

            /* 修飾キーは event 優先、取得できなければ keyboardState にフォールバック / Read modifiers from the event first, fall back to keyboardState */
            var shiftPressed = event.shiftKey;
            if (shiftPressed === undefined) { shiftPressed = ScriptUI.environment.keyboardState.shiftKey; }
            var delta = shiftPressed ? 10 : 1;

            if (event.keyName === "Up") {
                /* Shift時は10の倍数にスナップ / Snap to multiples of 10 with Shift */
                value = shiftPressed ? Math.ceil((value + 1) / delta) * delta : value + delta;
                event.preventDefault();
            } else if (event.keyName === "Down") {
                value = shiftPressed ? Math.floor((value - 1) / delta) * delta : value - delta;
                if (value < 0) { value = 0; }
                event.preventDefault();
            } else {
                return;
            }

            editText.text = Math.round(value);
            /* 連動している依存フィールドを更新する / Fire onChanging so linked fields update */
            if (typeof editText.onChanging === "function") { editText.onChanging(); }
        });
    }

    /* 前後の空白を除去する（ES3にString.trimが無いため） / Trim whitespace (ES3 has no String.trim) */
    function trim(text) {
        return ("" + text).replace(/^\s+/, "").replace(/\s+$/, "");
    }

    /**
     * 「3, 4」「3-5」形式の指定を0始まりのアートボード索引配列に厳密変換する
     * Strictly parse a "3, 4" / "3-5" style spec into 0-based artboard indices
     * 各トークンを正規表現で完全一致検証し、"3abc"・"1-2-3"・空トークン(",")等は不正扱い。
     * @param {string} text 入力文字列（1始まり）
     * @param {number} count アートボード総数
     * @returns {array} 索引配列。不正な場合は null
     */
    function parseArtboardSelection(text, count) {
        var singlePattern = /^\d+$/;                 /* 単一番号 / single number */
        var rangePattern = /^\d+\s*-\s*\d+$/;        /* 範囲（前後の空白可） / range (spaces allowed) */
        var indices = [];
        var seen = {};
        var tokens = ("" + text).split(",");
        for (var i = 0; i < tokens.length; i++) {
            var token = trim(tokens[i]);
            /* 空トークン（"1," ",2" "1,,2" 等）は不正扱い / empty token (from "1," ",2" "1,,2") is invalid */
            if (token === "") { return null; }

            var start, end;
            if (rangePattern.test(token)) {
                var dash = token.indexOf("-");
                start = parseInt(trim(token.substring(0, dash)), 10);
                end = parseInt(trim(token.substring(dash + 1)), 10);
            } else if (singlePattern.test(token)) {
                start = end = parseInt(token, 10);
            } else {
                return null; /* 形式不一致（"3abc" "1-2-3" 等） / format mismatch */
            }
            if (start > end) { var swap = start; start = end; end = swap; } /* 逆順は入れ替え / swap reversed range */

            for (var n = start; n <= end; n++) {
                if (n < 1 || n > count) { return null; } /* 範囲外は不正 / out of range is invalid */
                var idx = n - 1;
                if (!seen[idx]) { seen[idx] = true; indices.push(idx); } /* 重複は1回だけ / dedupe */
            }
        }
        return indices.length ? indices : null;
    }

    /**
     * 対象パネル（現在のアートボード／すべてのアートボード／指定）を構築する
     * Build the target panel (Current artboard / All artboards / Specify)
     * @param {Window} dialog 追加先ダイアログ
     * @param {string} defaultSelection 指定入力欄の初期値（例: "1-6"）
     */
    function addTargetPanel(dialog, defaultSelection) {
        var panel = dialog.add("panel", undefined, getLocalizedText("panel.target"));
        setupPanel(panel, 6);

        /* 「現在のアートボード」ラジオ / "Current artboard" radio */
        var currentRow = panel.add("group");
        currentRow.orientation = "row";
        currentRow.alignment = "left";
        var currentRadio = currentRow.add("radiobutton", undefined, getLocalizedText("radio.current"));

        /* 「すべてのアートボード」ラジオ / "All artboards" radio */
        var allRow = panel.add("group");
        allRow.orientation = "row";
        allRow.alignment = "left";
        var allRadio = allRow.add("radiobutton", undefined, getLocalizedText("radio.all"));

        /* 「指定」ラジオ＋範囲入力 / "Specify" radio with range input */
        var specifyRow = panel.add("group");
        specifyRow.orientation = "row";
        specifyRow.alignment = "left";
        var specifyRadio = specifyRow.add("radiobutton", undefined, getLocalizedText("radio.specify"));
        var selectInput = specifyRow.add("edittext", undefined, defaultSelection);
        selectInput.characters = 8; /* 通常のフィールドの2倍幅 / Twice the usual field width */

        var radios = [currentRadio, allRadio, specifyRadio];

        /* ラジオは親が異なると排他にならないため手動で同期 / Sync manually since radios in different parents are not exclusive */
        function selectTarget(active) {
            for (var i = 0; i < radios.length; i++) { radios[i].value = (radios[i] === active); }
            selectInput.enabled = (active === specifyRadio); /* 入力欄は「指定」時のみ有効 / input enabled only for "Specify" */
            requestPreview(dialog);
        }
        currentRadio.onClick = function () { selectTarget(currentRadio); };
        allRadio.onClick = function () { selectTarget(allRadio); };
        specifyRadio.onClick = function () { selectTarget(specifyRadio); };
        selectInput.onChanging = function () { requestPreview(dialog); };

        /* 初期状態は「すべてのアートボード」 / Default to "All artboards" */
        allRadio.value = true;
        selectInput.enabled = false;

        dialog.currentRadio = currentRadio;
        dialog.allRadio = allRadio;
        dialog.specifyRadio = specifyRadio;
        dialog.selectInput = selectInput;
    }

    /**
     * UI明度(0..1)を取得する（失敗時は1=明るい） / Get UI brightness (0..1); 1 on failure
     * @returns {number} 0..1 の明度
     */
    function getUIBrightness() {
        try {
            var brightness = app.preferences.getRealPreference("uiBrightness");
            if (brightness < 0) { brightness = 0; }
            if (brightness > 1) { brightness = 1; }
            return brightness;
        } catch (e) {
            return 1;
        }
    }

    /* 明るいUIかどうか / Whether the UI is light */
    function isLightUI() {
        return getUIBrightness() > 0.5;
    }

    /* 基準点セルの配色（initAnchorColorsでUI明暗に合わせて上書き） / Anchor-cell colors (overwritten by initAnchorColors) */
    var ANCHOR_LINE_COLOR = [0.6, 0.6, 0.6, 1];      /* 枠・ケイ線：薄いグレー / border & rules: light gray */
    var ANCHOR_SELECTED_FILL = [0.4, 0.4, 0.4, 1];   /* 選択セルの塗り / fill of the selected cell */

    /* UI明暗に合わせて選択セルの塗り色を決める（表示前に呼ぶ） / Decide the selected-cell fill from the light/dark UI (call before showing) */
    function initAnchorColors() {
        ANCHOR_SELECTED_FILL = isLightUI() ? [0.4, 0.4, 0.4, 1] : [0.8, 0.8, 0.8, 1];
    }

    /* コントロールを再描画（notifyは環境により例外を投げ得るので保護） / Redraw a control (notify can throw in some environments) */
    function redrawControl(control) {
        try { control.notify("onDraw"); } catch (e) {}
    }

    /* 正方形のサブパスを1つ作る / Build one square subpath */
    function squarePath(graphics, x, y, size) {
        graphics.newPath();
        graphics.moveTo(x, y);
        graphics.lineTo(x + size, y);
        graphics.lineTo(x + size, y + size);
        graphics.lineTo(x, y + size);
        graphics.closePath();
    }

    /* 基準点セルの□を1つ描画（選択時のみ塗り、枠は常時） / Draw one anchor-cell square (fill only when selected, always bordered) */
    function drawAnchorCell(graphics, x, y, size, selected) {
        if (selected) {
            squarePath(graphics, x, y, size);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, ANCHOR_SELECTED_FILL));
        }
        /* 枠線は常に薄いグレー（塗りの上に描く） / Border always light gray (drawn over the fill) */
        squarePath(graphics, x, y, size);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, ANCHOR_LINE_COLOR, 1));
    }

    /* 9軸ウィジェットを描画（外周の□をケイ線でつなぐ・中央は独立） / Draw the 9-axis widget (outer squares joined by rules; center stands alone) */
    function drawAnchorWidget(widget) {
        var graphics = widget.graphics;
        var width = widget.size[0];
        var height = widget.size[1];

        try {
            /* コントロール地色で塗って透過に見せる / Paint the control's own background so it looks transparent */
            graphics.rectPath(0, 0, width, height);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e0) {}

        var cellSize = 9;
        var cellGap = 7.5;
        var cellStep = cellSize + cellGap;
        var gridSize = cellSize * 3 + cellGap * 2;
        var originX = Math.round((width - gridSize) / 2);
        var originY = Math.round((height - gridSize) / 2);

        function anchorCellX(index) { return originX + (index % 3) * cellStep; }
        function anchorCellY(index) { return originY + Math.floor(index / 3) * cellStep; }

        /* 中央(4)を除く外周の□どうしをケイ線でつなぐ / Join the outer squares (except center 4) with rules */
        var connections = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];
        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, ANCHOR_LINE_COLOR, 1);
        for (var i = 0; i < connections.length; i++) {
            var cellA = connections[i][0];
            var cellB = connections[i][1];
            var cellAX = anchorCellX(cellA);
            var cellAY = anchorCellY(cellA);
            var cellBX = anchorCellX(cellB);
            var cellBY = anchorCellY(cellB);
            graphics.newPath();
            if (cellB - cellA === 1) {
                graphics.moveTo(cellAX + cellSize, cellAY + cellSize / 2);
                graphics.lineTo(cellBX, cellBY + cellSize / 2);
            } else {
                graphics.moveTo(cellAX + cellSize / 2, cellAY + cellSize);
                graphics.lineTo(cellBX + cellSize / 2, cellBY);
            }
            graphics.strokePath(linePen);
        }

        for (var index = 0; index < 9; index++) {
            drawAnchorCell(graphics, anchorCellX(index), anchorCellY(index), cellSize, index === widget.anchorIndex);
        }
    }

    /**
     * 3×3の基準点ウィジェット（onDrawで描画するproxy）を構築する / Build a 3×3 anchor proxy widget drawn via onDraw
     * 選択に応じて dialog.anchorX / dialog.anchorY に 0/0.5/1 の割合を設定する（既定=左上）。
     * @param {Window} dialog 割合を書き込むダイアログ
     * @param {Group} parent 追加先グループ
     */
    function addAnchorGrid(dialog, parent) {
        var column = parent.add("group");
        column.orientation = "column";
        column.alignChildren = ["center", "top"];
        column.spacing = 4;
        column.add("statictext", undefined, getLocalizedText("label.anchor"));

        var widget = column.add("button", undefined, "");
        widget.preferredSize = [66, 66];
        widget.minimumSize = [66, 66];
        widget.maximumSize = [66, 66];
        widget.anchorIndex = 0; /* 0..8 行優先（0=左上, 4=中央, 8=右下）/ 0..8 row-major (0=top-left, 4=center, 8=bottom-right) */
        widget.onDraw = function () { drawAnchorWidget(this); };
        widget.onClick = function () {}; /* セル判定は mousedown 側 / hit-testing happens in mousedown */

        /* 選択索引を割合(0/0.5/1)に変換して dialog に反映 / Map the index to fractions and store on the dialog */
        function commitAnchor() {
            dialog.anchorX = (widget.anchorIndex % 3) * 0.5;
            dialog.anchorY = Math.floor(widget.anchorIndex / 3) * 0.5;
        }
        commitAnchor(); /* 既定=左上 / default: top-left */

        /* クリックした3×3のセルを基準点に設定（座標はコントロール基準）/ Set the anchor from the clicked 3x3 cell (control-relative coords) */
        try {
            widget.addEventListener("mousedown", function (event) {
                var col = Math.floor(event.clientX / (widget.size[0] / 3));
                var row = Math.floor(event.clientY / (widget.size[1] / 3));
                if (col < 0) { col = 0; }
                if (col > 2) { col = 2; }
                if (row < 0) { row = 0; }
                if (row > 2) { row = 2; }
                widget.anchorIndex = row * 3 + col;
                commitAnchor();
                redrawControl(widget);
                requestPreview(dialog);
            });
        } catch (e) {}
    }

    /**
     * サイズ・スケール統合パネルを構築する / Build the merged size & scale panel
     * スケール%を一意の倍率とし、幅・高さ（現在の定規単位）はアクティブアートボードの現在サイズ基準で相互連動する。
     * The scale % is a single uniform ratio; width/height (in the current ruler unit) are linked to the active artboard's current size.
     * @param {Window} dialog 追加先ダイアログ
     * @param {string} unitLabel 幅・高さの単位ラベル（例: "mm"）
     * @param {number} baseWidth 幅の基準値：アクティブアートボードの現在幅を現在の定規単位へ変換済み（ptではない） / active artboard current width, converted to the current ruler unit (not pt)
     * @param {number} baseHeight 高さの基準値：アクティブアートボードの現在高さを現在の定規単位へ変換済み（ptではない） / active artboard current height, converted to the current ruler unit (not pt)
     */
    function addScalePanel(dialog, unitLabel, baseWidth, baseHeight) {
        var panel = dialog.add("panel", undefined, getLocalizedText("panel.scale"));
        setupPanel(panel, 6);

        /* 2列レイアウト：左=スケール/幅/高さの3行、右=基準点グリッド / Two columns: left has scale/width/height rows, right holds the anchor grid */
        var gridRow = panel.add("group");
        gridRow.orientation = "row";
        gridRow.alignChildren = ["left", "center"];
        gridRow.spacing = 16;

        var fieldColumn = gridRow.add("group");
        fieldColumn.orientation = "column";
        fieldColumn.alignChildren = ["left", "top"];
        fieldColumn.spacing = 6;
        var scaleInput = addSizeField(fieldColumn, labelText("field.scale"), "100", "%");
        var widthInput = addSizeField(fieldColumn, labelText("field.width"), formatNumber(baseWidth), unitLabel);
        var heightInput = addSizeField(fieldColumn, labelText("field.height"), formatNumber(baseHeight), unitLabel);

        /* 基準点グリッド（2列目） / Anchor reference-point grid (second column) */
        addAnchorGrid(dialog, gridRow);

        /* チェックボックス群（上に10pxの余白） / Checkbox group (10px top margin) */
        var checkboxGroup = panel.add("group");
        checkboxGroup.orientation = "column";
        checkboxGroup.alignChildren = ["left", "top"];
        checkboxGroup.margins = [0, 10, 0, 0];

        /* オブジェクトも一緒に拡大・縮小するか / Whether to scale the objects along with the artboard */
        var scaleObjectsCheckbox = checkboxGroup.add("checkbox", undefined, getLocalizedText("checkbox.scaleObjects"));
        scaleObjectsCheckbox.value = true;

        /* アートボードのX/Y/W/Hを整数化してピクセルグリッドに合わせる / Round artboard X/Y/W/H to integers for the pixel grid */
        var pixelGridCheckbox = checkboxGroup.add("checkbox", undefined, getLocalizedText("checkbox.pixelGrid"));
        pixelGridCheckbox.value = false;

        /* スケール%から幅・高さ(現在サイズ×%)を再計算する / Recalc width/height (current size × %) from the scale */
        function applyScaleToSize() {
            var percent = parseFloat(scaleInput.text);
            if (isNaN(percent)) { return; }
            widthInput.text = formatNumber(baseWidth * percent / 100);
            heightInput.text = formatNumber(baseHeight * percent / 100);
        }

        /**
         * 編集中の寸法欄からスケール%を逆算し、スケール欄ともう一方の寸法欄だけ更新する
         * Derive the scale from the edited size field, updating only the scale field and the OTHER size field
         * （編集中の欄自身は書き換えない＝小数点入力が消える不具合を防ぐ / never rewrite the field being edited, so decimals can be typed）
         * @param {number} editedValue 編集中の欄の値
         * @param {number} editedBase 編集中の欄の基準サイズ
         * @param {EditText} otherField もう一方（連動更新する）寸法欄
         * @param {number} otherBase もう一方の基準サイズ
         */
        function applySizeToScale(editedValue, editedBase, otherField, otherBase) {
            if (isNaN(editedValue) || editedBase === 0) { return; }
            var percent = editedValue / editedBase * 100;
            scaleInput.text = formatNumber(percent);
            otherField.text = formatNumber(otherBase * percent / 100);
        }

        /* 直近の有効なスケール%（入力強化の復帰先） / Last valid scale % (fallback for input hardening) */
        var lastValidScale = 100;

        /* 確定時に3欄をスケール基準へ正規化する（不正値は直近の有効値へ復帰） / On commit, canonicalize all three fields to the scale (invalid → last valid) */
        function normalizeFields() {
            var percent = parseFloat(scaleInput.text);
            if (isNaN(percent) || percent <= 0) {
                percent = lastValidScale; /* 空・0・負・非数値は直近の有効値へ / empty/0/negative/NaN falls back */
            } else {
                lastValidScale = percent;
            }
            scaleInput.text = formatNumber(percent);
            widthInput.text = formatNumber(baseWidth * percent / 100);
            heightInput.text = formatNumber(baseHeight * percent / 100);
            requestPreview(dialog);
        }

        scaleInput.onChanging = function () {
            applyScaleToSize();
            requestPreview(dialog);
        };
        widthInput.onChanging = function () {
            applySizeToScale(parseFloat(widthInput.text), baseWidth, heightInput, baseHeight);
            requestPreview(dialog);
        };
        heightInput.onChanging = function () {
            applySizeToScale(parseFloat(heightInput.text), baseHeight, widthInput, baseWidth);
            requestPreview(dialog);
        };
        /* 確定（Enter/フォーカスアウト）で正規化 / Normalize on commit (Enter / focus-out) */
        scaleInput.onChange = normalizeFields;
        widthInput.onChange = normalizeFields;
        heightInput.onChange = normalizeFields;
        scaleObjectsCheckbox.onClick = function () {
            requestPreview(dialog);
        };
        pixelGridCheckbox.onClick = function () {
            requestPreview(dialog);
        };

        dialog.scaleInput = scaleInput;
        dialog.widthInput = widthInput;
        dialog.heightInput = heightInput;
        dialog.scaleObjectsCheckbox = scaleObjectsCheckbox;
        dialog.pixelGridCheckbox = pixelGridCheckbox;
    }

    /**
     * サイズ入力ダイアログを構築する / Build the size-input dialog
     * @param {string} unitLabel 表示する単位ラベル（例: "mm"）
     * @param {string} defaultWidth 幅入力欄の初期値
     * @param {string} defaultHeight 高さ入力欄の初期値
     * @param {number} artboardCount アートボード総数（選択の初期値に使用）
     * @returns {Window} ダイアログウィンドウ（各入力コントロールを公開）
     */
    function createResizeDialog(unitLabel, defaultWidth, defaultHeight, artboardCount) {
        /* 基準点ウィジェットの配色をUI明暗に合わせる / Match the anchor widget colors to the light/dark UI */
        initAnchorColors();

        /* タイトルバーにバージョンを表示 / Show version in the title bar */
        var dialog = new Window("dialog", getLocalizedText("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        /* 対象パネル / Target panel */
        addTargetPanel(dialog, "1-" + artboardCount);

        /* サイズ・スケール統合パネル（アクティブアートボードの現在サイズを基準に） / Merged size & scale panel (based on the active artboard's current size) */
        addScalePanel(dialog, unitLabel, parseFloat(defaultWidth), parseFloat(defaultHeight));

        /* 数値入力欄に↑↓キーでの増減を付与 / Enable arrow-key value change on the numeric fields */
        changeValueByArrowKey(dialog.scaleInput);
        changeValueByArrowKey(dialog.widthInput);
        changeValueByArrowKey(dialog.heightInput);

        /* ボタン類はパネル幅いっぱいには広げず右寄せ / Keep buttons right-aligned, not full width */
        var btnGroup = dialog.add("group");
        btnGroup.orientation = "row";
        btnGroup.alignment = "right";
        btnGroup.add("button", undefined, getLocalizedText("button.cancel"), {name: "cancel"});
        btnGroup.add("button", undefined, getLocalizedText("button.apply"), {name: "ok"});

        return dialog;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* プレビュー更新を要求する（コントローラ未接続なら何もしない） / Request a preview refresh (no-op until wired) */
    function requestPreview(dialog) {
        if (dialog.onPreview) { dialog.onPreview(); }
    }

    /* 1アイテムを基準点基準に逆拡縮して元へ戻す（rollback用） / Inverse-scale one item about the anchor to undo it (for rollback) */
    function inverseResizeItem(item, ratioX, ratioY, anchorX, anchorY, originalPosition) {
        try {
            var inverseX = (1 / ratioX) * 100;
            var inverseY = (1 / ratioY) * 100;
            /* 第7引数(changeLineWidths)は線幅拡縮のパーセント値を渡す（参考実装準拠） / 7th arg is the line-width scale percentage (per the reference) */
            item.resize(inverseX, inverseY, true, true, true, true, inverseX, Transformation.TOPLEFT);
            item.position = [originalPosition[0], originalPosition[1]];
        } catch (e) {}
    }

    /**
     * 【確定用】指定比率でオブジェクト群を基準点(anchor)基準に一発で拡縮する（線幅も拡縮／途中失敗は自前で完全復元）
     * [For commit] Scale items about the anchor once (also scales line width); on mid-way failure, fully reverts its own items
     * artboardsResizeWithObjects.jsx(Alexander Ladygin) を参考にした resize()+position 方式。
     * resize() でサイズ・線幅を拡縮（アイテム左上基準）し、position で基準点からのオフセットを拡縮して再配置する。
     * オブジェクトごとに { item, originalPosition, resized } を記録し、失敗時はこの呼び出しで変形済みの分を確実に戻す。
     * @returns {object} { success:boolean, transformedItems:array[, error:Error] }
     */
    function resizeItemsAbout(items, ratioX, ratioY, anchorX, anchorY) {
        var transformedItems = [];
        if (!items || !items.length) { return { success: true, transformedItems: transformedItems }; }
        var states = []; /* {item, originalPosition, resized} 途中失敗時の復元用 / for rollback on failure */
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            var originalPosition = item.position; // [左, 上]（ドキュメント座標） / [left, top] in document coordinates
            var state = { item: item, originalPosition: [originalPosition[0], originalPosition[1]], resized: false };
            states.push(state);
            try {
                /* サイズ・線幅を拡縮（アイテム左上基準）。第7引数=線幅拡縮のパーセント値（参考実装準拠） / Scale size & line width about the item's top-left; 7th arg = line-width scale percentage (per the reference) */
                item.resize(ratioX * 100, ratioY * 100, true, true, true, true, ratioX * 100, Transformation.TOPLEFT);
                state.resized = true; /* resize成功。以降のpositionで失敗しても逆resizeで戻せる / resize done; still recoverable if position fails */
                /* 基準点からのオフセットを拡縮して再配置 / Reposition by scaling the offset from the anchor point */
                item.position = [
                    anchorX + (originalPosition[0] - anchorX) * ratioX,
                    anchorY + (originalPosition[1] - anchorY) * ratioY
                ];
                transformedItems.push(item);
            } catch (e) {
                /* 途中失敗：この呼び出しで resize 済み（position前後どちらも）を確実に復元 / mid-way failure: revert every item resized in this call */
                for (var r = states.length - 1; r >= 0; r--) {
                    if (states[r].resized) {
                        inverseResizeItem(states[r].item, ratioX, ratioY, anchorX, anchorY, states[r].originalPosition);
                    }
                }
                return { success: false, transformedItems: [], error: e };
            }
        }
        return { success: true, transformedItems: transformedItems };
    }

    /**
     * 指定アートボード上の編集可能オブジェクトを安全に取得する（selection/active が途中例外でも壊れない）
     * Safely collect editable objects on the given artboard (selection/active stay consistent even if an exception occurs)
     */
    function getObjectsOnArtboard(doc, index) {
        var items = [];
        try {
            doc.selection = null;
            doc.artboards.setActiveArtboardIndex(index);
            doc.selectObjectsOnActiveArtboard();
            var selection = doc.selection;
            /* selection が null/未定義相当や配列でない場合も安全に扱う / Handle null / non-array selection safely */
            if (selection && typeof selection.length === "number") {
                for (var i = 0; i < selection.length; i++) { items.push(selection[i]); }
            }
        } finally {
            /* 例外の有無に関わらず選択を解除する / clear the selection whether or not an exception occurred */
            try { doc.selection = null; } catch (e) {}
        }
        return items;
    }

    /* オブジェクトの一意識別子を返す（uuid優先、無ければ null） / Return an object's unique id (prefer uuid; null if unavailable) */
    function getItemUuid(item) {
        try {
            if (item.uuid) { return item.uuid; }
        } catch (e) {}
        return null;
    }

    /* 点(x,y)がアートボード矩形[左,上,右,下]の内側か / Whether point (x,y) is inside the artboard rect [L,T,R,B] */
    function rectContainsPoint(rect, x, y) {
        return x >= rect[0] && x <= rect[2] && y <= rect[1] && y >= rect[3];
    }

    /* オブジェクト境界[左,上,右,下]とアートボード矩形の重なり面積 / Overlap area between object bounds [L,T,R,B] and an artboard rect */
    function overlapArea(bounds, rect) {
        var overlapWidth = Math.min(bounds[2], rect[2]) - Math.max(bounds[0], rect[0]);
        var overlapHeight = Math.min(bounds[1], rect[1]) - Math.max(bounds[3], rect[3]);
        return (overlapWidth > 0 && overlapHeight > 0) ? overlapWidth * overlapHeight : 0;
    }

    /**
     * オブジェクトの所属アートボードを決める（中心点包含→重なり最大→番号が小さい方）
     * Decide which artboard owns an object (center containment → max overlap → smallest index)
     * @param {array} bounds オブジェクト境界 [左,上,右,下]
     * @param {array} sortedTargets 昇順の対象アートボード索引
     * @param {object} rectByIndex index→矩形
     * @returns {number} 所属アートボード索引（どこにも重ならなければ -1）
     */
    function pickOwnerArtboard(bounds, sortedTargets, rectByIndex) {
        var centerX = (bounds[0] + bounds[2]) / 2;
        var centerY = (bounds[1] + bounds[3]) / 2;
        /* 1. 中心点が含まれるアートボード（昇順で最初） / center-point containment (first in ascending order) */
        for (var i = 0; i < sortedTargets.length; i++) {
            if (rectContainsPoint(rectByIndex[sortedTargets[i]], centerX, centerY)) { return sortedTargets[i]; }
        }
        /* 2/3. 重なり面積が最大のもの（同じなら昇順の先勝ち＝番号が小さい方） / largest overlap (ties break to the smallest index via ascending scan) */
        var bestIndex = -1;
        var bestArea = 0;
        for (var j = 0; j < sortedTargets.length; j++) {
            var area = overlapArea(bounds, rectByIndex[sortedTargets[j]]);
            if (area > bestArea) { bestArea = area; bestIndex = sortedTargets[j]; }
        }
        return bestArea > 0 ? bestIndex : -1;
    }

    /**
     * 対象アートボード群から、重複を排除して所属先ごとにオブジェクトをまとめる
     * Collect objects across target artboards, deduped and grouped by owning artboard
     * @returns {object} index→[所属オブジェクト] / index → owned items
     */
    function collectOwnedObjectsByArtboard(doc, sortedTargets, rectByIndex) {
        var owners = {};
        for (var t = 0; t < sortedTargets.length; t++) { owners[sortedTargets[t]] = []; }

        /* 重複判定はオブジェクト参照で行う（uuid優先、無ければ === 比較） / Dedupe by object identity (uuid first; === fallback) */
        var seenUuids = {};
        var seenItems = [];
        function alreadySeen(item) {
            var uuid = getItemUuid(item);
            if (uuid !== null) {
                if (seenUuids[uuid]) { return true; }
                seenUuids[uuid] = true;
                return false;
            }
            for (var k = 0; k < seenItems.length; k++) {
                if (seenItems[k] === item) { return true; } /* 同一参照なら重複 / same reference = duplicate */
            }
            seenItems.push(item);
            return false;
        }

        for (var s = 0; s < sortedTargets.length; s++) {
            var items = getObjectsOnArtboard(doc, sortedTargets[s]);
            for (var i = 0; i < items.length; i++) {
                if (alreadySeen(items[i])) { continue; }
                var owner = pickOwnerArtboard(items[i].geometricBounds, sortedTargets, rectByIndex);
                if (owner !== -1) { owners[owner].push(items[i]); }
            }
        }
        return owners;
    }

    /**
     * 基準点を固定したまま新しいアートボード矩形を算出する（ピクセルグリッド時は幅高さと基準点を整数化）
     * Compute the new artboard rect keeping the anchor fixed (integerize size & pivot when pixel-grid is on)
     *
     * ピクセルグリッド仕様 / Pixel-grid behavior:
     *   基準点固定と「幅・高さの整数化」を優先する。基準点も整数化するため、
     *   基準点でない側の端（右端／下端など）は 0.5px 単位になる場合がある。
     *   Anchor-fixed and integer width/height take priority; since the pivot is also integerized,
     *   the non-anchored edges (e.g. right/bottom) may land on 0.5px.
     *
     * effectiveRatioX/Y は整数化後の実効倍率（オブジェクト変形はこれを使い、アートボードと一致させる）。
     * effectiveRatioX/Y are the post-rounding effective ratios; object scaling uses them so objects match the artboard.
     * @returns {object} { rect:[左,上,右,下], pivotX, pivotY, effectiveRatioX, effectiveRatioY }
     */
    function computeNewGeometry(rect, settings) {
        var width = rect[2] - rect[0];
        var height = rect[1] - rect[3];
        /* 1. 元の矩形から基準点（固定される点）を計算 / pivot (fixed point) from the original rect */
        var pivotX = rect[0] + settings.anchorFx * width;
        var pivotY = rect[1] - settings.anchorFy * height;

        /* 2. 新しい幅・高さ / new width & height */
        var newWidth = width * settings.ratioX;
        var newHeight = height * settings.ratioY;

        if (settings.pixelGrid) {
            /* 3. 幅・高さを整数化（0以下防止） / integerize size (prevent <= 0) */
            newWidth = Math.round(newWidth);
            newHeight = Math.round(newHeight);
            if (newWidth < 1) { newWidth = 1; }
            if (newHeight < 1) { newHeight = 1; }
            /* 4. 基準点も整数化 / integerize the pivot too */
            pivotX = Math.round(pivotX);
            pivotY = Math.round(pivotY);
        }

        /* 5. 基準点を固定して矩形を再計算（Y座標は下方向がマイナス） / rebuild the rect keeping the pivot fixed (Y grows downward negative) */
        var newLeft = pivotX - settings.anchorFx * newWidth;
        var newTop = pivotY + settings.anchorFy * newHeight;
        return {
            rect: [newLeft, newTop, newLeft + newWidth, newTop - newHeight],
            pivotX: pivotX,
            pivotY: pivotY,
            /* 整数化後の実効倍率（幅・高さが0のケースは1倍扱い） / post-rounding effective ratio (treat 0 size as 1x) */
            effectiveRatioX: width !== 0 ? newWidth / width : 1,
            effectiveRatioY: height !== 0 ? newHeight / height : 1
        };
    }

    /* ダイアログの対象指定から処理するアートボード索引配列を得る（不正は null） / Resolve target artboard indices (null if invalid) */
    function resolveTargetIndices(dialog, count) {
        if (dialog.currentRadio.value) {
            /* 現在（起動時）のアクティブアートボードのみ / only the active artboard at launch */
            return [dialog.currentArtboardIndex];
        }
        if (dialog.allRadio.value) {
            var all = [];
            for (var i = 0; i < count; i++) { all.push(i); }
            return all;
        }
        return parseArtboardSelection(dialog.selectInput.text, count);
    }

    /* ダイアログの現在値を読み取る（不正なら null） / Read the current dialog settings (null if invalid) */
    function readSettings(dialog, count) {
        var indices = resolveTargetIndices(dialog, count);
        if (!indices) { return null; }
        var scale = parseFloat(dialog.scaleInput.text);
        if (isNaN(scale) || scale <= 0) { return null; }
        var ratio = scale / 100; /* スケールは一意（縦横同率） / Uniform scale (same ratio for both axes) */
        return {
            indices: indices,
            ratioX: ratio,
            ratioY: ratio,
            anchorFx: dialog.anchorX, /* 0=左,0.5=中央,1=右 / 0=left,0.5=center,1=right */
            anchorFy: dialog.anchorY, /* 0=上,0.5=中央,1=下 / 0=top,0.5=center,1=bottom */
            scaleObjects: dialog.scaleObjectsCheckbox.value,
            pixelGrid: dialog.pixelGridCheckbox.value
        };
    }

    /**
     * コントローラを生成する / Create the controller
     * プレビューはアートボード矩形のみを更新（完全可逆・線幅に無関係）。オブジェクトはOK確定時に一度だけ resize() する。
     * The live preview updates artboard rectangles only (fully reversible, unrelated to line width);
     * objects are scaled once with resize() at commit. Initial state is the baseline for restore/re-apply to limit drift.
     * オブジェクトは複数アートボードにまたがっても所属ルールで一意化し1回だけ変形する。
     * @param {Document} doc 対象ドキュメント
     * @param {Window} dialog 設定を読むダイアログ
     * @param {number} count アートボード総数
     * @returns {object} { update, commit, restore, restoreSelectionAndActive, hasError }
     */
    function createController(doc, dialog, count) {
        /* 初期状態を保存（キャンセル・確定時に復元） / Save initial state (restored on cancel/commit) */
        var originalActiveIndex = doc.artboards.getActiveArtboardIndex();
        var originalSelection = [];
        var initialSelection = doc.selection;
        if (initialSelection && typeof initialSelection.length === "number") {
            for (var s = 0; s < initialSelection.length; s++) { originalSelection.push(initialSelection[s]); }
        }

        var capturedRects = {};  /* index -> [元rect] / index -> original rect */
        var lastError = null;    /* 直近のエラー / last error */

        /* アートボードの元rectを一度だけ保存する / Capture an artboard's original rect once */
        function captureRect(index) {
            if (!capturedRects[index]) {
                var rect = doc.artboards[index].artboardRect;
                capturedRects[index] = [rect[0], rect[1], rect[2], rect[3]];
            }
            return capturedRects[index];
        }

        /**
         * アートボード矩形を初期スナップショットへ戻す（成功可否を返す） / Reset artboard rects to the initial snapshot (returns success)
         * @returns {object} { success:boolean[, error:Error] }
         */
        function restore() {
            try {
                for (var key in capturedRects) {
                    if (!capturedRects.hasOwnProperty(key)) { continue; }
                    doc.artboards[Number(key)].artboardRect = capturedRects[key];
                }
                return { success: true };
            } catch (e) {
                /* 復元失敗時は履歴（capturedRects）を消さない / keep the history (capturedRects) on failure */
                return { success: false, error: e };
            }
        }

        /* 元の選択状態とアクティブアートボードを可能な範囲で復元する / Restore the original selection and active artboard as far as possible */
        function restoreSelectionAndActive() {
            try { doc.selection = null; } catch (e0) {}
            for (var i = 0; i < originalSelection.length; i++) {
                /* 削除・無効化された項目は個別にスキップ / skip deleted/invalid items individually */
                try { originalSelection[i].selected = true; } catch (e1) {}
            }
            try { doc.artboards.setActiveArtboardIndex(originalActiveIndex); } catch (e2) {}
        }

        /* 昇順の対象索引配列を返す（所属ルールのタイブレーク用） / Return target indices sorted ascending (for ownership tie-breaks) */
        function sortedTargetsOf(settings) {
            var sorted = settings.indices.concat();
            sorted.sort(function (a, b) { return a - b; });
            return sorted;
        }

        /* プレビュー：対象アートボードの矩形だけを更新する（オブジェクトは触らない） / Preview: update only the target artboard rectangles (objects untouched) */
        function update() {
            lastError = null;
            try {
                restore(); /* まず矩形を初期状態へ / reset rects to the initial state first */

                var settings = readSettings(dialog, count);
                if (!settings) { app.redraw(); return; } /* 不正入力中は何も適用しない / apply nothing while input is invalid */

                var sortedTargets = sortedTargetsOf(settings);
                for (var t = 0; t < sortedTargets.length; t++) {
                    var index = sortedTargets[t];
                    captureRect(index);
                    doc.artboards[index].artboardRect = computeNewGeometry(capturedRects[index], settings).rect;
                }
            } catch (e) {
                lastError = e;
                restore();
            }
            app.redraw();
        }

        /**
         * 確定：矩形を初期状態へ戻し、初期状態から一度だけ「矩形＋オブジェクト」を適用する（線幅も正しく拡縮・累積誤差なし）
         * Commit: reset rects, then apply "rects + objects" once from the initial state (correct line width, no drift)
         */
        function commit() {
            lastError = null;
            /* position/geometricBounds を artboardRect と同じドキュメント座標で扱うため明示設定（終了時に復元） / Force document coordinates so position/geometricBounds match artboardRect (restored at the end) */
            var previousCoordinateSystem = app.coordinateSystem;
            app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
            try {
                restore(); /* プレビューの矩形変更を初期状態へ戻す / reset the preview rects to the initial state */

                var settings = readSettings(dialog, count);
                if (!settings) { app.redraw(); return; }

                var sortedTargets = sortedTargetsOf(settings);
                var t;
                for (t = 0; t < sortedTargets.length; t++) { captureRect(sortedTargets[t]); }
                var ownedObjects = settings.scaleObjects
                    ? collectOwnedObjectsByArtboard(doc, sortedTargets, capturedRects)
                    : {};

                var committed = []; /* 失敗時のベストエフォート巻き戻し用 / for best-effort rollback on failure */
                for (t = 0; t < sortedTargets.length; t++) {
                    var index = sortedTargets[t];
                    var geometry = computeNewGeometry(capturedRects[index], settings);

                    /* オブジェクトはアートボードと同じ基準点・実効倍率（整数化後）で拡縮 / Objects use the artboard's pivot and post-rounding effective ratio */
                    if (settings.scaleObjects) {
                        var owned = ownedObjects[index] || [];
                        var result = resizeItemsAbout(owned, geometry.effectiveRatioX, geometry.effectiveRatioY, geometry.pivotX, geometry.pivotY);
                        committed.push({ items: result.transformedItems, ratioX: geometry.effectiveRatioX, ratioY: geometry.effectiveRatioY, pivotX: geometry.pivotX, pivotY: geometry.pivotY });
                        if (!result.success) {
                            /* 途中失敗：確定済みを逆拡縮し、矩形を戻して中止 / mid-way failure: inverse-resize committed items, reset rects, and stop */
                            lastError = result.error;
                            for (var c = committed.length - 1; c >= 0; c--) {
                                resizeItemsAbout(committed[c].items, 1 / committed[c].ratioX, 1 / committed[c].ratioY, committed[c].pivotX, committed[c].pivotY);
                            }
                            restore();
                            app.redraw();
                            return;
                        }
                    }

                    doc.artboards[index].artboardRect = geometry.rect;
                }
            } catch (e) {
                lastError = e;
                restore();
            } finally {
                /* 座標系を元へ戻す / restore the coordinate system */
                try { app.coordinateSystem = previousCoordinateSystem; } catch (e2) {}
            }
            app.redraw();
        }

        return {
            update: update,
            commit: commit,
            restore: restore,
            restoreSelectionAndActive: restoreSelectionAndActive,
            hasError: function () { return lastError !== null; }
        };
    }

    /* エントリポイント：ダイアログを出し、アートボードサイズをプレビューしつつ確定でオブジェクトも拡縮する / Entry point: show dialog, preview artboard size, scale objects on commit */
    function resizeArtboards() {
        if (app.documents.length === 0) {
            alert(getLocalizedText("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var unit = getRulerUnit();
        var count = doc.artboards.length;

        /* アクティブアートボードの現在サイズを初期値にする / Use the active artboard's current size as defaults */
        var activeRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        var defaultWidth = formatNumber((activeRect[2] - activeRect[0]) / unit.factor);
        var defaultHeight = formatNumber((activeRect[1] - activeRect[3]) / unit.factor);

        var dialog = createResizeDialog(unit.label, defaultWidth, defaultHeight, count);
        /* 「現在のアートボード」対象用に起動時のアクティブ索引を保持 / Remember the launch-time active index for the "Current artboard" target */
        dialog.currentArtboardIndex = doc.artboards.getActiveArtboardIndex();

        /* プレビュー配線（矩形のみ）＋表示時にスケール欄へフォーカス / Wire up the preview (rects only) and focus the scale field on show */
        var controller = createController(doc, dialog, count);
        dialog.onPreview = controller.update;
        dialog.onShow = function () {
            controller.update();
            dialog.scaleInput.active = true;
        };

        var result = dialog.show();

        if (result !== 1) {
            /* キャンセル：矩形を復元（失敗は通知）し、選択・アクティブアートボードを戻す / Cancel: restore rects (notify on failure), restore selection & active artboard */
            var cancelRestore = controller.restore();
            controller.restoreSelectionAndActive();
            app.redraw();
            if (!cancelRestore.success) { alert(getLocalizedText("alert.restoreError")); }
            return;
        }

        /* OK：不正値は矩形を戻してエラー表示 / OK: revert rects and report on invalid input */
        if (!readSettings(dialog, count)) {
            controller.restore();
            controller.restoreSelectionAndActive();
            app.redraw();
            alert(resolveTargetIndices(dialog, count) ? getLocalizedText("alert.invalidNumber") : getLocalizedText("alert.invalidSelection"));
            return;
        }

        /* 確定：矩形を戻し、resize()で「矩形＋オブジェクト」を一度だけ適用 / Commit: reset rects, then apply rects + objects once via resize() */
        controller.commit();
        if (controller.hasError()) {
            /* 変形失敗時は確定せず、矩形と選択・アクティブを復元 / on failure, do not commit; restore rects, selection & active */
            controller.restore();
            controller.restoreSelectionAndActive();
            app.redraw();
            alert(getLocalizedText("alert.transformError"));
            return;
        }

        /* 適用済みの状態を確定。選択・アクティブアートボードのみ復元（完了メッセージは表示しない） / Keep the applied result; restore only selection & active artboard (no completion message) */
        controller.restoreSelectionAndActive();
        app.redraw();
    }

    resizeArtboards();

})();
