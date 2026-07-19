#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
GenerateGuidesGrid.jsx

### 概要

- Illustrator のアートボード、または選択オブジェクトの外接矩形を指定した行数・列数に分割し、グリッド用のガイドを生成するスクリプトです。
- セルの長方形化（不透明度はスライダーで調整・初期15%）、角丸（ライブエフェクト）、各セルの左右分割ガイド、中心点の表示、プリセット書き出しに対応しています。
- 対象（選択オブジェクト／アートボード／すべてのアートボード）をパネルで切り替えられ、選択オブジェクト基準のときは元オブジェクトの処理（削除する／そのまま／ガイド化）を選べます。
- 入力変更（キーボード・矢印キー・チェックボックス・スライダー）でプレビューが即時更新され、Undo で巻き戻してから再描画するため履歴を汚しにくい構成です。
- 主要なコントロールにはツールチップ（操作説明）を表示します。

### 主な機能

- 行数・列数、行間・列間ガター、上下左右マージン、ガイドの伸張距離を設定
- ガター連動（行間に連動）／マージン連動（上の値を左右下に反映、上下左右が異なるプリセット選択時は自動でOFF）
- セルの長方形化／不透明度スライダー（0〜100%）／角丸（ライブエフェクト）／中心点の表示
- 各セルを左右分割（各セルの左右中央に縦ガイドを作成）
- ガイド描画ON/OFF（OFF時は伸張設定を連動して無効化）／既存ガイドの削除
- アウトライン⇔プレビュー表示の切り替えボタン
- 現在設定をプリセットとして書き出し
- 単一／すべてのアートボードに適用（単一時は「すべてのアートボード」を無効化）
- 選択オブジェクトの外接矩形を対象にグリッド生成（複数選択・任意形状対応）
- 元のオブジェクトを削除する／そのまま維持する／ガイド化する処理
- マージン・ガター過大時は描画をスキップして不正な長方形生成を防止
- 日本語／英語インターフェース自動切替（ツールチップも日英対応）

### オリジナル、謝辞

スガサワ君β / https://note.com/sgswkn/n/nee8c3ec1a14c

### 更新履歴

- v1.7.0 (20260612) : 角丸（ライブエフェクト）・中心点の表示・各セルの左右分割ガイド・不透明度スライダーを追加、元オブジェクト処理に「ガイド化」を追加（「隠す」は廃止）、アウトライン⇔プレビュー切り替えボタンとツールチップを追加、UI レイアウト・文言を整理（マージンは3×3配置）、裁ち落とし機能を削除、単一アートボード時は「すべてのアートボード」を無効化、過大マージン/ガターの安全チェックを追加、内部リファクタ（重複削減・命名整理・drawGrid の分離）
- v1.6.6 (20260316) : 起動時の対象モードを自動判定（選択あり→選択オブジェクト、なし→アートボード）
- v1.0 (20250424) : 初期バージョン
*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GenerateGuidesGrid";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

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

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;

    // =========================================
    // ローカライズ / Localization
    // =========================================
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

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
            linkMargin: { ja: "連動", en: "Same Value" },
            cellRect: { ja: "長方形化", en: "Rectangles" },
            showCenter: { ja: "中心点を表示", en: "Show Center Point" },
            roundCorner: { ja: "角丸", en: "Round Corners" },
            splitCell: { ja: "各セルを左右分割", en: "Split Each Cell Horizontally" },
            drawGuides: { ja: "ガイドを引く", en: "Draw Guides" },
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
            linkGutter: { ja: "列間を行間と同じ値に保ちます。", en: "Keep the column gutter equal to the row gutter." },
            linkMargin: { ja: "上の値を下・左・右にも適用します。", en: "Apply the top value to bottom, left, and right." },
            cellRect: { ja: "各セルを長方形として作成します。", en: "Create a rectangle for each cell." },
            showCenter: { ja: "作成した長方形の中心点を属性パネルで表示します（OK時に適用）。", en: "Show the center point of created rectangles in the Attributes panel (applied on OK)." },
            roundCorner: { ja: "各長方形に角丸（ライブエフェクト）を適用します。", en: "Apply the Round Corners live effect to each rectangle." },
            opacity: { ja: "長方形の不透明度（0〜100%）。", en: "Opacity of the rectangles (0–100%)." },
            splitCell: { ja: "各セルの左右中央に縦ガイドを作成します。", en: "Create a vertical guide at each cell's horizontal center." },
            guideExtension: { ja: "アートボード／オブジェクトの外側へガイドを伸ばす距離。", en: "Distance to extend guides beyond the artboard/object." },
            clearGuides: { ja: "描画前に専用レイヤー（grid_guides）内の既存ガイドを削除します。", en: "Remove existing guides in the dedicated layer (grid_guides) before drawing." },
            toGuide: { ja: "元の選択オブジェクトをガイドに変換します。", en: "Convert the original selected objects into guides." },
            outline: { ja: "アウトライン表示とプレビュー表示を切り替えます。", en: "Toggle between Outline and Preview view." },
            exportPreset: { ja: "現在の設定をプリセットファイルに書き出します。", en: "Export the current settings to a preset file." }
        },
        /* ボタン / Buttons */
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            apply: { ja: "適用", en: "Apply" },
            ok: { ja: "OK", en: "OK" },
            outline: { ja: "アウトライン", en: "Outline" },
            preview: { ja: "プレビュー", en: "Preview" },
            exportPreset: { ja: "書き出し", en: "Export" }
        }
    };

    /* ラベル取得（ドット区切りキー、{slash} は "/" に置換）/ Get label (dotted key, {slash} replaced with "/") */
    function getLabel(key) {
        var parts = key.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (node && node[parts[i]] !== undefined) {
                node = node[parts[i]];
            } else {
                return key;
            }
        }
        var text = (node && node[currentLanguage] !== undefined) ? node[currentLanguage] : key;
        return String(text).replace(/\{slash\}/g, "/");
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(key) {
        return getLabel(key) + (currentLanguage === "ja" ? "：" : ":");
    }

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* グループの共通設定（row/column で整列を切り替え）/ Apply shared group layout (alignChildren switches by orientation) */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* 見出しに単位を付与（日本語は全角括弧、英語は半角括弧）/ Append unit to a title (full-width parens JA, half-width EN) */
    function titleWithUnit(key) {
        return getLabel(key) + (currentLanguage === "ja" ? "（" + unitLabel + "）" : " (" + unitLabel + ")");
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
    var PRESET_UNIT_SCALE = 0.5; /* 換算時の控えめ係数（pt/px のときは等倍）/ Modest coefficient when converting (1:1 for pt/px) */
    /* pt/px（unitFactor が 1）のときは等倍、それ以外へ換算するときだけ係数を掛ける / Scale only when actually converting to a non-point unit */
    function presetScale() {
        return (unitFactor === 1.0) ? 1.0 : PRESET_UNIT_SCALE;
    }
    /* pt → 現在単位（必要時のみ係数を掛けて整数に丸め）/ points → current unit (scaled when converting, rounded to an integer) */
    function ptToUnit(ptValue) {
        return Math.round(Number(ptValue) / unitFactor * presetScale());
    }
    /* 現在単位 → pt（係数で割り戻して整数に丸め）/ current unit → points (inverse scale, rounded to an integer) */
    function unitToPt(value) {
        return Math.round(Number(value) * unitFactor / presetScale());
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
            colGutter: 0,
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
            colGutter: 0,
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
            colGutter: 0,
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
            colGutter: 0,
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
            colGutter: 50,
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
            colGutter: 30,
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
            colGutter: 20,
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
            colGutter: 20,
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
            colGutter: 0,
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
            colGutter: 0,
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
            colGutter: 0,
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
    }

    // プレビュー手順を1つ実行してカウント / Run an action as a preview step and count it
    PreviewManager.prototype.addStep = function (func) {
        try {
            func();
            this.undoDepth++;
            app.redraw();
        } catch (e) {
            alert("Preview Error: " + e);
        }
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
    // 描画ヘルパー（doc を受け取り、UI/クロージャに依存しない）/ Drawing helpers (take doc; no UI/closure deps)
    // =========================================

    // 関数を実行し、例外を握りつぶす（try/catch の重複を集約）/ Run a function, swallowing errors (centralizes the try/catch pattern)
    function safeExecute(fn) {
        try {
            fn();
        } catch (e) {
            $.writeln("[GenerateGuidesGrid] safeExecute error: " + e);
        }
    }

    // src の自前プロパティを target にコピー / Copy own properties from src into target
    function extend(target, src) {
        for (var k in src) {
            if (src.hasOwnProperty(k)) target[k] = src[k];
        }
        return target;
    }

    // レイヤーのロックを安全に解除 / Safely unlock a layer
    function safeUnlockLayer(layer) {
        safeExecute(function () {
            if (layer && layer.locked) layer.locked = false;
        });
    }

    // 指定名のレイヤーを安全に削除 / Safely remove a layer by name
    function safeRemoveLayerByName(doc, name) {
        safeExecute(function () {
            var layer = doc.layers.getByName(name);
            if (layer) layer.remove();
        });
    }

    // 指定名のレイヤーを取得、なければ作成 / Get a layer by name, creating it if missing
    function getOrCreateLayer(doc, name) {
        var layer;
        try {
            layer = doc.layers.getByName(name);
        } catch (e) {
            layer = doc.layers.add();
            layer.name = name;
        }
        return layer;
    }

    // ガイド線を1本追加（塗り・線なしのパスをガイド化してレイヤー先頭へ）/ Add one guide line
    function addGuideLine(doc, layer, startPoint, endPoint) {
        var guideLine = doc.pathItems.add();
        guideLine.setEntirePath([startPoint, endPoint]);
        guideLine.stroked = false;
        guideLine.filled = false;
        guideLine.guides = true;
        guideLine.move(layer, ElementPlacement.PLACEATBEGINNING);
        return guideLine;
    }

    // 角丸ライブエフェクトのXML（radius は pt）/ Round Corners live-effect XML (radius in pt)
    function roundCornersEffectXML(radius) {
        var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius #value# "/></LiveEffect>';
        return xml.replace('#value#', radius);
    }

    // 黒色作成（CMYK／RGB対応）/ Create black color (CMYK/RGB)
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

    // グリッド（ガイド＋セル長方形）を描画し、作成したセル長方形の配列を返す / Draw the grid; return created cell rects
    function drawGrid(ctx) {
        var doc = ctx.doc;
        var isPreview = ctx.isPreview;
        var columnCount = ctx.columnCount;
        var rowCount = ctx.rowCount;
        if (isNaN(columnCount) || columnCount <= 0 || isNaN(rowCount) || rowCount <= 0) return [];

        var ext = ctx.ext;
        var top = ctx.marginTop, bottom = ctx.marginBottom, left = ctx.marginLeft, right = ctx.marginRight;
        var rowGutter = ctx.rowGutter, colGutter = ctx.colGutter;
        var drawCells = ctx.drawCells, drawGuidesNow = ctx.drawGuidesNow, cornerRadius = ctx.cornerRadius;
        var splitCells = ctx.splitCells;
        var cellOpacity = (typeof ctx.cellOpacity === "number") ? ctx.cellOpacity : CELL_OPACITY;
        var createdCells = [];

        // プレビューは1枚に全部描く。確定時はガイド用レイヤーはガイド系を描くときだけ作る
        // Preview draws everything on one layer; on commit, create the guide layer only when guides are drawn
        var gridLayerName = isPreview ? PREVIEW_LAYER_NAME : GUIDE_LAYER_NAME;
        var needGridLayer = isPreview ? (drawGuidesNow || splitCells || drawCells) : (drawGuidesNow || splitCells);
        var gridLayer = needGridLayer ? getOrCreateLayer(doc, gridLayerName) : null;
        safeUnlockLayer(gridLayer);

        var cellLayer = gridLayer; // プレビューはセルも gridLayer に描く / preview: cells live on gridLayer
        if (!isPreview && drawCells) {
            cellLayer = getOrCreateLayer(doc, CELL_LAYER_NAME);
            safeUnlockLayer(cellLayer);
        }

        var targetRects = [];
        if (ctx.selBounds) {
            targetRects.push(ctx.selBounds);
        } else {
            for (var b = 0; b < doc.artboards.length; b++) {
                if (!ctx.allBoards && b !== doc.artboards.getActiveArtboardIndex()) continue;
                targetRects.push(doc.artboards[b].artboardRect);
            }
        }

        for (var ti = 0; ti < targetRects.length; ti++) {
            var targetRect = targetRects[ti];
            var abLeft = targetRect[0],
                abTop = targetRect[1],
                abRight = targetRect[2],
                abBottom = targetRect[3];
            var baseLeft = abLeft + left;
            var baseRight = abRight - right;
            var baseTop = abTop - top;
            var baseBottom = abBottom + bottom;

            var usableWidth = baseRight - baseLeft;
            var usableHeight = baseTop - baseBottom;
            var totalColGutter = (columnCount - 1) * colGutter;
            var totalRowGutter = (rowCount - 1) * rowGutter;
            // マージン・ガターが過大でセル幅/高さが0以下になる場合はこの対象をスキップ
            // Skip this target if margins/gutters are too large (cell width/height would be <= 0)
            if (usableWidth - totalColGutter <= 0 || usableHeight - totalRowGutter <= 0) continue;
            var cellWidth = (usableWidth - totalColGutter) / columnCount;
            var cellHeight = (usableHeight - totalRowGutter) / rowCount;

            var guideLeft = abLeft - ext;
            var guideRight = abRight + ext;
            var guideTop = abTop + ext;
            var guideBottom = abBottom - ext;

            if (drawGuidesNow) {
                if (columnCount === 1 && rowCount === 1) {
                    // 四辺（マージン適用後の有効領域）をガイド化 / Four edges of the usable area
                    addGuideLine(doc, gridLayer, [guideLeft, baseTop], [guideRight, baseTop]);
                    addGuideLine(doc, gridLayer, [guideLeft, baseBottom], [guideRight, baseBottom]);
                    addGuideLine(doc, gridLayer, [baseLeft, guideTop], [baseLeft, guideBottom]);
                    addGuideLine(doc, gridLayer, [baseRight, guideTop], [baseRight, guideBottom]);
                } else {
                    // 通常ガイド描画（行・列）/ Normal grid guides (rows and columns)
                    var y = baseTop;
                    addGuideLine(doc, gridLayer, [guideLeft, y], [guideRight, y]);
                    for (var j = 0; j < rowCount; j++) {
                        y -= cellHeight;
                        addGuideLine(doc, gridLayer, [guideLeft, y], [guideRight, y]);
                        if (j < rowCount - 1) {
                            y -= rowGutter;
                            addGuideLine(doc, gridLayer, [guideLeft, y], [guideRight, y]);
                        }
                    }
                    addGuideLine(doc, gridLayer, [guideLeft, baseBottom], [guideRight, baseBottom]);

                    var x = baseLeft;
                    addGuideLine(doc, gridLayer, [x, guideTop], [x, guideBottom]);
                    for (var i2 = 0; i2 < columnCount; i2++) {
                        x += cellWidth;
                        addGuideLine(doc, gridLayer, [x, guideTop], [x, guideBottom]);
                        if (i2 < columnCount - 1) {
                            x += colGutter;
                            addGuideLine(doc, gridLayer, [x, guideTop], [x, guideBottom]);
                        }
                    }
                    addGuideLine(doc, gridLayer, [baseRight, guideTop], [baseRight, guideBottom]);
                }
            }

            if (drawCells && cellLayer) {
                var startX = baseLeft;
                var startY = baseTop;
                for (var row = 0; row < rowCount; row++) {
                    var cellY = startY - (cellHeight + rowGutter) * row;
                    for (var col = 0; col < columnCount; col++) {
                        var cellX = startX + (cellWidth + colGutter) * col;
                        var cellRect = cellLayer.pathItems.rectangle(cellY, cellX, cellWidth, cellHeight);
                        cellRect.stroked = false;
                        cellRect.filled = true;
                        cellRect.fillColor = createBlackColor(doc);
                        cellRect.opacity = cellOpacity;
                        if (cornerRadius > 0) cellRect.applyEffect(roundCornersEffectXML(cornerRadius));
                        createdCells.push(cellRect);
                    }
                }
            }

            // 各セルを分割：セルの左右中央に縦ガイド / Split each cell: vertical guide at each cell's horizontal center
            if (splitCells && gridLayer) {
                var splitStartX = baseLeft;
                var splitStartY = baseTop;
                for (var sr = 0; sr < rowCount; sr++) {
                    var splitCellTop = splitStartY - (cellHeight + rowGutter) * sr;
                    var splitCellBottom = splitCellTop - cellHeight;
                    for (var sc = 0; sc < columnCount; sc++) {
                        var splitCenterX = splitStartX + (cellWidth + colGutter) * sc + cellWidth / 2;
                        addGuideLine(doc, gridLayer, [splitCenterX, splitCellTop], [splitCenterX, splitCellBottom]);
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
        return createdCells;
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
        var previewMgr = new PreviewManager();

        // 確定描画したセル長方形（中心点表示の対象）/ Cell rects from final draw (targets for center-point display)
        var drawnCellItems = [];

        // 選択オブジェクトの参照と外接矩形をキャッシュ / Cache selection refs and bounds before dialog
        var cachedSelectionItems = [];
        var cachedSelectionBounds = (function () {
            var sel = doc.selection;
            if (!sel || sel.length === 0) return null;
            for (var si = 0; si < sel.length; si++) {
                cachedSelectionItems.push(sel[si]);
            }
            var firstBounds = sel[0].geometricBounds;
            var sLeft = firstBounds[0], sTop = firstBounds[1], sRight = firstBounds[2], sBottom = firstBounds[3];
            for (var s = 1; s < sel.length; s++) {
                var itemBounds = sel[s].geometricBounds;
                if (itemBounds[0] < sLeft) sLeft = itemBounds[0];
                if (itemBounds[1] > sTop) sTop = itemBounds[1];
                if (itemBounds[2] > sRight) sRight = itemBounds[2];
                if (itemBounds[3] < sBottom) sBottom = itemBounds[3];
            }
            return [sLeft, sTop, sRight, sBottom];
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
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.spacing = 10;

        // プリセット選択＋書き出し / Preset selection and export
        var presetGroup = dialog.add("group");
        setupGroup(presetGroup, "row");
        presetGroup.alignment = ["center", "top"]; // 左右中央 / horizontally centered
        presetGroup.margins = [0, 0, 0, 10]; // 下に少し余白 / add some bottom margin

        presetGroup.add("statictext", undefined, labelText("field.preset"));
        var presetDropdown = presetGroup.add("dropdownlist", undefined, []);
        presetDropdown.selection = 0;
        var exportPresetButton = presetGroup.add("button", undefined, getLabel("button.exportPreset"));
        exportPresetButton.helpTip = getLabel("tooltip.exportPreset");

        // プリセット書き出し / Export the current settings as a preset
        exportPresetButton.onClick = function () {
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
                columns: parseInt(columnCountInput.text, 10),
                rows: parseInt(rowCountInput.text, 10),
                guideExtension: unitToPt(extensionInput.text),
                marginTop: unitToPt(inputTop.text),
                marginBottom: unitToPt(inputBottom.text),
                marginLeft: unitToPt(inputLeft.text),
                marginRight: unitToPt(inputRight.text),
                rowGutter: unitToPt(inputRowGutter.text),
                colGutter: unitToPt(inputColGutter.text),
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
                'colGutter: ' + currentPreset.colGutter + ', ' +
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
        setupGroup(targetRow, "row", 12);
        targetRow.alignChildren = ["left", "fill"];

        // 対象パネル：対象の種類（縦並び）/ Target panel: target type (vertical)
        var targetPanel = targetRow.add("panel", undefined, getLabel("panel.target"));
        setupPanel(targetPanel);
        var targetTypeGroup = targetPanel.add("group");
        setupGroup(targetTypeGroup, "column");
        var rbTargetSelection = targetTypeGroup.add("radiobutton", undefined, getLabel("target.selection"));
        var rbTargetArtboard = targetTypeGroup.add("radiobutton", undefined, getLabel("target.artboard"));
        var rbTargetAllArtboards = targetTypeGroup.add("radiobutton", undefined, getLabel("target.allArtboards"));

        // 元のオブジェクトパネル（対象パネルの外・横並び、選択オブジェクト時のみ有効）/ Original-object panel (outside target panel, selection mode only)
        var objActionGroup = targetRow.add("panel", undefined, getLabel("panel.originalObject"));
        setupPanel(objActionGroup);
        var objActionRow = objActionGroup.add("group");
        setupGroup(objActionRow, "column");
        var removeOriginalRadio = objActionRow.add("radiobutton", undefined, getLabel("radio.remove"));
        var keepOriginalRadio = objActionRow.add("radiobutton", undefined, getLabel("radio.keep"));
        var makeOriginalGuidesRadio = objActionRow.add("radiobutton", undefined, getLabel("radio.toGuide"));
        makeOriginalGuidesRadio.helpTip = getLabel("tooltip.toGuide");
        keepOriginalRadio.value = true; // デフォルト / default

        // 選択がなければ「選択オブジェクト」を無効化 / Disable selection option when nothing is selected
        if (!cachedSelectionBounds) {
            rbTargetSelection.enabled = false;
        }
        // アートボードが1つだけなら「すべてのアートボード」を無効化 / Disable "All Artboards" when only one artboard
        if (!hasMultipleArtboards) {
            rbTargetAllArtboards.enabled = false;
        }
        // 初期選択 / Initial target radio
        rbTargetSelection.value = (targetMode === "selection");
        rbTargetArtboard.value = (targetMode === "artboard");
        rbTargetAllArtboards.value = (targetMode === "allArtboards");

        // 対象モード変更 / Target mode change
        function onTargetChange() {
            if (rbTargetSelection.value) {
                targetMode = "selection";
            } else if (rbTargetAllArtboards.value) {
                targetMode = "allArtboards";
            } else {
                targetMode = "artboard";
            }
            updateTargetMode();
            safeUpdatePreview();
        }
        rbTargetSelection.onClick = rbTargetArtboard.onClick = rbTargetAllArtboards.onClick = onTargetChange;

        // 元オブジェクト処理の切替でプレビュー更新 / Update preview on original-object radio change
        removeOriginalRadio.onClick = keepOriginalRadio.onClick = makeOriginalGuidesRadio.onClick = function () {
            safeUpdatePreview();
        };

        // 対象モードに応じた表示制御 / Enable/disable controls by target mode
        function updateTargetMode() {
            var selMode = isSelectionMode();
            objActionGroup.enabled = selMode;
        }

        // グリッド設定グループ / Grid settings group
        var rowColumnGroup = dialog.add("group");
        setupGroup(rowColumnGroup, "row", 12);
        rowColumnGroup.alignChildren = ["left", "top"];
        var gridLabelWidth = (currentLanguage === "ja") ? 40 : 50; // unify Number/Gutter label width and right-align

        // 行設定パネル / Row settings panel
        var rowPanel = rowColumnGroup.add("panel", undefined, getLabel("panel.row"));
        setupPanel(rowPanel);

        var rowCountGroup = rowPanel.add("group");
        setupGroup(rowCountGroup, "row");
        var lblRows = rowCountGroup.add("statictext", undefined, labelText("field.rowCount"));
        lblRows.justification = "right";
        lblRows.minimumSize.width = gridLabelWidth;
        lblRows.maximumSize.width = gridLabelWidth;
        var rowCountInput = rowCountGroup.add("edittext", undefined, "2");
        rowCountInput.characters = 3;

        var rowGutterGroup = rowPanel.add("group");
        setupGroup(rowGutterGroup, "row");
        var lblRowGutter = rowGutterGroup.add("statictext", undefined, labelText("field.rowGutter"));
        lblRowGutter.justification = "right";
        lblRowGutter.minimumSize.width = gridLabelWidth;
        lblRowGutter.maximumSize.width = gridLabelWidth;
        var inputRowGutter = rowGutterGroup.add("edittext", undefined, "0");
        inputRowGutter.characters = 3;
        rowGutterGroup.add("statictext", undefined, unitLabel);

        // 列設定パネル / Column settings panel
        var columnPanel = rowColumnGroup.add("panel", undefined, getLabel("panel.column"));
        setupPanel(columnPanel);

        var columnCountGroup = columnPanel.add("group");
        setupGroup(columnCountGroup, "row");
        var lblCols = columnCountGroup.add("statictext", undefined, labelText("field.columnCount"));
        lblCols.justification = "right";
        lblCols.minimumSize.width = gridLabelWidth;
        lblCols.maximumSize.width = gridLabelWidth;
        var columnCountInput = columnCountGroup.add("edittext", undefined, "2");
        columnCountInput.characters = 3;

        var colGutterGroup = columnPanel.add("group");
        setupGroup(colGutterGroup, "row");
        var lblColGutter = colGutterGroup.add("statictext", undefined, labelText("field.columnGutter"));
        lblColGutter.justification = "right";
        lblColGutter.minimumSize.width = gridLabelWidth;
        lblColGutter.maximumSize.width = gridLabelWidth;
        var inputColGutter = colGutterGroup.add("edittext", undefined, "0");
        inputColGutter.characters = 3;
        colGutterGroup.add("statictext", undefined, unitLabel);

        // 行間に連動（列パネル下部）/ Link to row gutter (under column panel)
        var linkGutterCheckbox = columnPanel.add("checkbox", undefined, getLabel("checkbox.linkGutter"));
        linkGutterCheckbox.helpTip = getLabel("tooltip.linkGutter");
        linkGutterCheckbox.value = true;

        // 連動ON時は列間を行間と同じ値にする / When linked, sync col gutter to row gutter
        function syncGutterLink() {
            if (linkGutterCheckbox.value) {
                inputColGutter.text = inputRowGutter.text;
                colGutterGroup.enabled = false;
            } else {
                var xVal = parseInt(columnCountInput.text, 10);
                colGutterGroup.enabled = (xVal > 1);
            }
        }

        linkGutterCheckbox.onClick = function () {
            syncGutterLink();
            safeUpdatePreview();
        };

        // マージン全体パネル / Margin panel
        var marginPanel = dialog.add("panel", undefined, titleWithUnit("panel.margin"));
        setupPanel(marginPanel);
        // 3×3 グリッド配置（中央=連動）/ 3×3 grid layout (center = link)
        var MARGIN_CELL_WIDTH = (currentLanguage === "ja") ? 78 : 92;

        // ラベル＋数値のセル / A label+field cell
        function addMarginCell(parentRow, labelKey) {
            var cell = parentRow.add("group");
            cell.orientation = "row";
            cell.alignment = ["center", "center"];
            cell.minimumSize.width = MARGIN_CELL_WIDTH;
            cell.add("statictext", undefined, labelText(labelKey));
            var input = cell.add("edittext", undefined, "0");
            input.characters = 3;
            return { group: cell, input: input };
        }
        // 位置合わせ用の空セル / Empty cell for alignment
        function addMarginSpacer(parentRow) {
            var cell = parentRow.add("group");
            cell.minimumSize.width = MARGIN_CELL_WIDTH;
        }

        // 1行目：［空］［上］［空］/ Row 1: [empty][top][empty]
        var marginRow1 = marginPanel.add("group");
        marginRow1.orientation = "row";
        marginRow1.alignment = ["center", "top"];
        addMarginSpacer(marginRow1);
        var topCell = addMarginCell(marginRow1, "field.top");
        var inputTop = topCell.input;
        addMarginSpacer(marginRow1);

        // 2行目：［左］［連動］［右］/ Row 2: [left][link][right]
        var marginRow2 = marginPanel.add("group");
        marginRow2.orientation = "row";
        marginRow2.alignment = ["center", "top"];
        var leftCell = addMarginCell(marginRow2, "field.left");
        var marginLeftGroup = leftCell.group;
        var inputLeft = leftCell.input;
        var marginCommonGroup = marginRow2.add("group");
        marginCommonGroup.orientation = "row";
        marginCommonGroup.alignment = ["center", "center"];
        marginCommonGroup.minimumSize.width = MARGIN_CELL_WIDTH;
        var commonMarginCheckbox = marginCommonGroup.add("checkbox", undefined, getLabel("checkbox.linkMargin"));
        commonMarginCheckbox.helpTip = getLabel("tooltip.linkMargin");
        commonMarginCheckbox.value = true; // デフォルトでON / on by default
        var rightCell = addMarginCell(marginRow2, "field.right");
        var marginRightGroup = rightCell.group;
        var inputRight = rightCell.input;

        // 3行目：［空］［下］［空］/ Row 3: [empty][bottom][empty]
        var marginRow3 = marginPanel.add("group");
        marginRow3.orientation = "row";
        marginRow3.alignment = ["center", "top"];
        addMarginSpacer(marginRow3);
        var bottomCell = addMarginCell(marginRow3, "field.bottom");
        var marginBottomGroup = bottomCell.group;
        var inputBottom = bottomCell.input;
        addMarginSpacer(marginRow3);

        // セルパネル＋ガイドパネルを横並び（左：セル、右：ガイド）/ Cell panel + Guides panel side by side (left: cell, right: guides)
        var guideRow = dialog.add("group");
        setupGroup(guideRow, "row", 12);
        guideRow.alignChildren = ["left", "fill"];

        // セルパネル（ガイドパネルの左）/ Cell panel (left of Guides panel)
        var cellPanel = guideRow.add("panel", undefined, getLabel("panel.options"));
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
        var inputRoundCorner = roundCornerGroup.add("edittext", undefined, "3");
        inputRoundCorner.characters = 2;
        roundCornerGroup.add("statictext", undefined, unitLabel);

        // 不透明度スライダー（0-100、数値表示なし）/ Opacity slider (0-100, no numeric readout)
        var opacityGroup = cellPanel.add("group");
        setupGroup(opacityGroup, "row");
        opacityGroup.add("statictext", undefined, labelText("field.opacity"));
        var opacitySlider = opacityGroup.add("slider", undefined, CELL_OPACITY, 0, 100);
        opacitySlider.helpTip = getLabel("tooltip.opacity");
        opacitySlider.preferredSize.width = 100;
        opacitySlider.onChanging = function () {
            safeUpdatePreview();
        };

        // ガイドパネル / Guides panel
        var guidesPanel = guideRow.add("panel", undefined, getLabel("panel.guides"));
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
        var extensionInput = extensionGroup.add("edittext", undefined, "10");
        extensionInput.characters = 2;
        extensionGroup.add("statictext", undefined, unitLabel);
        extensionCheckbox.onClick = function () {
            extensionInput.enabled = extensionCheckbox.value;
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
        clearGuidesCheckbox.value = true;

        // 数値フィールドを矢印キーで増減（Shift で10の倍数にスナップ）/ Adjust a numeric field with arrow keys (Shift snaps to multiples of 10)
        function changeValueByArrowKey(editText) {
            editText.addEventListener("keydown", function (event) {
                var value = Number(editText.text);
                if (isNaN(value)) return;

                var keyboard = ScriptUI.environment.keyboardState;

                if (event.keyName == "Up" || event.keyName == "Down") {
                    var isUp = event.keyName == "Up";
                    var delta = 1;

                    if (keyboard.shiftKey) {
                        // Shiftキー押下時は10の倍数にスナップ
                        value = Math.floor(value / 10) * 10;
                        delta = 10;
                    }

                    value += isUp ? delta : -delta;
                    if (value < 0) value = 0; // 必要なら下限チェック

                    event.preventDefault();
                    editText.text = value;
                    // 連動・ガター・列間同期をまとめて反映 / Apply linked margin / gutter enable / col gutter sync together
                    if (editText === inputTop && commonMarginCheckbox.value) {
                        syncCommonMargin();
                    }
                    if (editText === columnCountInput || editText === rowCountInput) {
                        updateGutterEnable();
                    }
                    if (editText === inputRowGutter && linkGutterCheckbox.value) {
                        inputColGutter.text = inputRowGutter.text;
                    }
                    // 入力変更を即時プレビューに反映 / Refresh preview immediately (Undo-safe)
                    safeUpdatePreview();
                }
            });
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
                previewMgr.rollback();
            } catch (e) {
                $.writeln("[GenerateGuidesGrid] preview rollback error: " + e);
            }

            // Draw preview as one undoable step
            previewMgr.addStep(function () {
                drawGrid(buildDrawContext(true)); // プレビュー描画 / draw as preview
            });

            // 選択オブジェクトの表示/非表示をプレビュー / Preview hide/show of selected objects
            if (isSelectionMode() && cachedSelectionItems.length > 0) {
                var shouldHide = (removeOriginalRadio && removeOriginalRadio.value);
                if (shouldHide) {
                    previewMgr.addStep(function () {
                        for (var ph = 0; ph < cachedSelectionItems.length; ph++) {
                            cachedSelectionItems[ph].hidden = true;
                        }
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

        // 各数値フィールドに矢印キー増減を付与 / Attach arrow-key adjustment to each numeric field
        changeValueByArrowKey(columnCountInput);
        changeValueByArrowKey(rowCountInput);
        changeValueByArrowKey(extensionInput);
        changeValueByArrowKey(inputTop);
        changeValueByArrowKey(inputBottom);
        changeValueByArrowKey(inputLeft);
        changeValueByArrowKey(inputRight);
        changeValueByArrowKey(inputRowGutter);
        changeValueByArrowKey(inputColGutter);

        // 入力中の変更もリアルタイム反映 / Attach onChanging for live preview
        attachLivePreview(columnCountInput);
        attachLivePreview(rowCountInput);
        attachLivePreview(extensionInput);
        // 上マージン変更時に連動ONなら左右下も同期 / Sync margins when top changes (if linked)
        inputTop.onChanging = function () {
            if (commonMarginCheckbox.value) {
                inputBottom.text = inputTop.text;
                inputLeft.text = inputTop.text;
                inputRight.text = inputTop.text;
            }
            safeUpdatePreview();
        };
        attachLivePreview(inputBottom);
        attachLivePreview(inputLeft);
        attachLivePreview(inputRight);
        // 行間変更時に連動チェックONなら列間も同期 / Sync col gutter when row gutter changes (if linked)
        inputRowGutter.onChanging = function () {
            if (linkGutterCheckbox.value) {
                inputColGutter.text = inputRowGutter.text;
            }
            safeUpdatePreview();
        };
        attachLivePreview(inputColGutter);

        // === ボタンエリア（3カラム：左アウトライン／中央スペーサー／右キャンセル・OK）/ Button area (3 columns: left outline / center spacer / right cancel+ok)
        var buttonArea = dialog.add("group");
        buttonArea.alignment = ["fill", "top"];
        buttonArea.orientation = "row";
        buttonArea.alignChildren = ["fill", "center"];
        buttonArea.margins = [0, 5, 0, 0]; // ボタンエリア上マージン +5 / extra top margin
        buttonArea.spacing = 0;

        // 左グループ（アウトラインボタン）/ Left group (Outline button)
        var buttonLeftGroup = buttonArea.add("group");
        buttonLeftGroup.orientation = "row";
        buttonLeftGroup.alignChildren = "left";
        var outlineButton = buttonLeftGroup.add("button", undefined, getLabel("button.outline"));
        outlineButton.helpTip = getLabel("tooltip.outline");
        // アウトライン⇔プレビュー表示を切り替え、ラベルもトグル / Toggle Outline/Preview view and the button label
        outlineButton.onClick = function () {
            app.executeMenuCommand('preview');
            outlineButton.text = (outlineButton.text === getLabel("button.outline"))
                ? getLabel("button.preview")
                : getLabel("button.outline");
        };

        // スペーサー（横に伸びる空白）/ Spacer (horizontal stretch)
        var spacer = buttonArea.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;
        spacer.maximumSize.height = 0;

        // 右グループ（キャンセル・OKボタン）/ Right group (Cancel/OK buttons)
        var buttonRightGroup = buttonArea.add("group");
        buttonRightGroup.alignment = ["right", "center"];
        buttonRightGroup.orientation = "row";
        buttonRightGroup.alignChildren = "right";
        buttonRightGroup.spacing = 10;
        var cancelButton = buttonRightGroup.add("button", undefined, getLabel("button.cancel"), {
            name: "cancel"
        });
        var okButton = buttonRightGroup.add("button", undefined, getLabel("button.ok"), {
            name: "ok"
        });
        okButton.alignment = ["right", "center"];

        // 表示用ラベルをローカライズ / Localize display label for dropdown
        function presetDisplayLabel(raw) {
            // 日本語UIのときは「 / 」以降を隠す / In Japanese UI, hide text after " / "
            if (currentLanguage === "ja") return String(raw).replace(/\s*\/.*$/, "");
            return raw;
        }

        // プリセットをドロップダウンに追加 / Add presets to dropdown
        for (var i = 0; i < presets.length; i++) {
            var display = presetDisplayLabel(presets[i].label);
            presetDropdown.add("item", display);
        }
        presetDropdown.selection = 0;

        // プリセットの値を入力欄に反映する共通関数 / Common function to apply preset values
        // 旧キー（x/y/top/bottom/left/right）も後方互換で受け付ける / Accept legacy keys for backward compatibility
        function applyPreset(preset) {
            /* 値を取得（新キー→旧キー→既定値の順）/ Pick a value: new key → legacy key → default */
            function pick(primary, legacy, dflt) {
                if (primary !== undefined) return primary;
                if (legacy !== undefined) return legacy;
                return dflt;
            }
            // 行数・列数（換算なし）/ Counts (no unit conversion)
            columnCountInput.text = pick(preset.columns, preset.x, 1);
            rowCountInput.text = pick(preset.rows, preset.y, 1);
            // 長さ系は pt 基準なので現在単位へ換算 / Length values are stored in pt — convert to the current unit
            extensionInput.text = ptToUnit(pick(preset.guideExtension, undefined, 0));
            var mTop = pick(preset.marginTop, preset.top, 0);
            var mBottom = pick(preset.marginBottom, preset.bottom, 0);
            var mLeft = pick(preset.marginLeft, preset.left, 0);
            var mRight = pick(preset.marginRight, preset.right, 0);
            inputTop.text = ptToUnit(mTop);
            inputBottom.text = ptToUnit(mBottom);
            inputLeft.text = ptToUnit(mLeft);
            inputRight.text = ptToUnit(mRight);
            inputRowGutter.text = ptToUnit(pick(preset.rowGutter, undefined, 0));
            inputColGutter.text = ptToUnit(pick(preset.colGutter, undefined, 0));
            // 上下左右が異なるプリセットは連動をOFF（連動が値を上書きして壊すのを防ぐ）
            // If margins differ, turn the link off so it won't overwrite the distinct values
            commonMarginCheckbox.value = (mTop === mBottom && mTop === mLeft && mTop === mRight);
            cellRectCheckbox.value = (typeof preset.drawCells !== "undefined") ? preset.drawCells : false;
            drawGuidesCheckbox.value = (typeof preset.drawGuides !== "undefined") ? preset.drawGuides : true;
            extensionGroup.enabled = drawGuidesCheckbox.value;
        }

        // プリセット選択時に入力値へ反映 / Apply preset values to inputs on selection
        presetDropdown.onChange = function () {
            applyPreset(presets[presetDropdown.selection.index]);
            updateGutterEnable();
            updateTargetMode();
            updateCenterEnable();
            syncCommonMargin();
            safeUpdatePreview();
        };

        // 初期プリセットの値を入力欄に反映 / Apply initial preset values to input fields
        applyPreset(presets[0]);

        // 「連動」同期処理 / Sync for "Link" margin
        function syncCommonMargin() {
            if (commonMarginCheckbox.value) {
                var val = inputTop.text;
                inputBottom.text = val;
                inputLeft.text = val;
                inputRight.text = val;
                marginBottomGroup.enabled = false;
                marginLeftGroup.enabled = false;
                marginRightGroup.enabled = false;
            } else {
                marginBottomGroup.enabled = true;
                marginLeftGroup.enabled = true;
                marginRightGroup.enabled = true;
            }
        }
        commonMarginCheckbox.onClick = function () {
            syncCommonMargin();
            safeUpdatePreview();
        };

        // ガター有効無効切り替え / Enable/disable gutter fields
        function updateGutterEnable() {
            var xVal = parseInt(columnCountInput.text, 10);
            var yVal = parseInt(rowCountInput.text, 10);
            rowGutterGroup.enabled = (yVal > 1);
            // 行数・列数が両方2以上のときのみ連動を有効 / Enable link only when both >= 2
            var bothMulti = (xVal > 1 && yVal > 1);
            linkGutterCheckbox.enabled = bothMulti;
            if (linkGutterCheckbox.value && bothMulti) {
                inputColGutter.text = inputRowGutter.text;
                colGutterGroup.enabled = false;
            } else {
                colGutterGroup.enabled = (xVal > 1);
            }
        }

        // 行数・列数変更時：ガター有効/無効を更新して即時プレビュー / On row/col change: refresh gutter enable + preview
        columnCountInput.onChanging = rowCountInput.onChanging = function () {
            updateGutterEnable();
            safeUpdatePreview();
        };

        // 「ガイドを引く」切替：伸張・クリアをディム制御してプレビュー / Toggle draw-guides: dim extension/clear, then preview
        drawGuidesCheckbox.onClick = function () {
            var enable = drawGuidesCheckbox.value;
            extensionGroup.enabled = enable;
            extensionInput.enabled = enable && extensionCheckbox.value;
            clearGuidesCheckbox.enabled = enable; // ガイドOFFならクリアをディム / dim clear when guides off
            safeUpdatePreview();
        };

        // 長方形化の有無でセル系オプション（中心点・角丸・不透明度）をディム制御 / Cell options dim unless rectangles are drawn
        // 角丸の数値欄は「長方形化ON かつ 角丸ON」のときだけ有効 / The round-corner field is enabled only when both rectangles and round corners are on
        function updateCenterEnable() {
            var cellsOn = cellRectCheckbox.value;
            centerPointCheckbox.enabled = cellsOn;
            roundCornerCheckbox.enabled = cellsOn;
            inputRoundCorner.enabled = cellsOn && roundCornerCheckbox.value;
            opacitySlider.enabled = cellsOn;
        }

        cellRectCheckbox.onClick = function () {
            updateCenterEnable();
            safeUpdatePreview();
        };

        // 角丸トグルで数値欄の有効/無効を更新してプレビュー / Toggle round corners: refresh the field enable, then preview
        roundCornerCheckbox.onClick = function () {
            updateCenterEnable();
            safeUpdatePreview();
        };
        changeValueByArrowKey(inputRoundCorner);
        attachLivePreview(inputRoundCorner);

        // OKボタン押下時（ドキュメントは変更せず閉じるだけ。クリア/描画は確定処理 finalAction で実行）
        // OK button: just close (no document mutation here — clear/draw happens in finalAction to keep undo bookkeeping correct)
        okButton.onClick = function () {
            updateGutterEnable();
            dialog.close(1);
        };

        // キャンセル：ダイアログを閉じるだけ（後処理は dialog.show() 後の分岐で rollback）/ Cancel: just close; cleanup/rollback happens after dialog.show()
        cancelButton.onClick = function () {
            dialog.close(0);
        };

        // 行・列・ガター・伸張・ガイド系の設定を収集 / Collect grid (rows/cols/gutters/extension/guide) settings
        function collectGridSettings() {
            return {
                columnCount: parseInt(columnCountInput.text, 10),
                rowCount: parseInt(rowCountInput.text, 10),
                ext: (extensionCheckbox.value ? parseFloat(extensionInput.text) : 0) * unitFactor,
                rowGutter: parseFloat(inputRowGutter.text) * unitFactor,
                colGutter: parseFloat(inputColGutter.text) * unitFactor,
                drawGuidesNow: drawGuidesCheckbox.value,
                splitCells: splitCellCheckbox.value
            };
        }

        // 上下左右マージンを収集 / Collect top/bottom/left/right margins
        function collectMarginSettings() {
            return {
                marginTop: parseFloat(inputTop.text) * unitFactor,
                marginBottom: parseFloat(inputBottom.text) * unitFactor,
                marginLeft: parseFloat(inputLeft.text) * unitFactor,
                marginRight: parseFloat(inputRight.text) * unitFactor
            };
        }

        // セル（長方形化・角丸・不透明度）の設定を収集 / Collect cell (rect/round/opacity) settings
        function collectCellSettings() {
            return {
                drawCells: cellRectCheckbox.value,
                cornerRadius: roundCornerCheckbox.value ? parseFloat(inputRoundCorner.text) * unitFactor : 0,
                cellOpacity: Math.round(opacitySlider.value)
            };
        }

        // 各 collector をまとめて描画コンテキストを構築 / Combine the collectors into a draw context
        function buildDrawContext(isPreview) {
            var ctx = {
                doc: doc,
                isPreview: isPreview,
                allBoards: isAllArtboards(),
                selBounds: isSelectionMode() ? cachedSelectionBounds : null
            };
            extend(ctx, collectGridSettings());
            extend(ctx, collectMarginSettings());
            extend(ctx, collectCellSettings());
            return ctx;
        }

        // 属性パネルの「中心を表示」を選択オブジェクトに適用 / Apply Attributes-panel "Show Center" to the current selection
        // API で直接設定できないため、一時アクション（.aia）を読み込んで再生 / No direct API, so load and play a temporary action
        function act_ShowCenter() {
            var actionSource = '/version 3' + '/name [ 9' + ' 417474726962757465' + ']' + '/isOpen 1' + '/actionCount 1' + '/action-1 {' + ' /name [ 10' + ' 53686f7743656e746572' + ' ]' + ' /keyIndex 0' + ' /colorIndex 0' + ' /isOpen 1' + ' /eventCount 1' + ' /event-1 {' + ' /useRulersIn1stQuadrant 0' + ' /internalName (adobe_attributePalette)' + ' /localizedName [ 12' + ' e5b19ee680a7e8a8ade5ae9a' + ' ]' + ' /isOpen 1' + ' /isOn 1' + ' /hasDialog 0' + ' /parameterCount 1' + ' /parameter-1 {' + ' /key 1668183154' + ' /showInPalette 4294967295' + ' /type (boolean)' + ' /value 1' + ' }' + ' }' + '}';

            // 他スクリプトと衝突しないよう temp 配下に固有名で書き出す / Write to temp with a script-specific name to avoid collisions
            var actionFile = new File(Folder.temp + "/GenerateGuidesGrid_ShowCenter.aia");
            if (!actionFile.open("w")) {
                return; // 書き込めなければ中止 / abort if it cannot be written
            }
            actionFile.write(actionSource);
            actionFile.close();

            app.loadAction(actionFile);
            actionFile.remove();
            // doScript が落ちても必ず unload する / Always unload, even if doScript throws
            try {
                app.doScript("ShowCenter", "Attribute", false); // action name, set name
            } finally {
                app.unloadAction("Attribute", ""); // set name
            }
        }

        // grid_guidesレイヤーのガイドだけ削除 / Remove only guides from grid_guides layer
        function clearExistingGuides() {
            var guidesLayer = null;
            for (var i3 = 0; i3 < doc.layers.length; i3++) {
                if (doc.layers[i3].name === GUIDE_LAYER_NAME) {
                    guidesLayer = doc.layers[i3];
                    break;
                }
            }
            if (guidesLayer) {
                safeUnlockLayer(guidesLayer);
                for (var j = guidesLayer.pageItems.length - 1; j >= 0; j--) {
                    var item = guidesLayer.pageItems[j];
                    if (item.guides) {
                        item.remove();
                    }
                }
            }
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
        updateGutterEnable();
        updateCenterEnable();
        syncCommonMargin();
        updateTargetMode();
        safeUpdatePreview();

        if (dialog.show() === 1) {
            // OK: rollback preview and execute final drawing once so user can undo in one step
            previewMgr.confirm(function () {
                // プレビューレイヤーを先に削除（最後に消すと選択が解除されるため）/ Remove preview layer first (removing it later clears the selection)
                safeRemoveLayerByName(doc, PREVIEW_LAYER_NAME);
                if (clearGuidesCheckbox.value) {
                    clearExistingGuides();
                }
                drawnCellItems = drawGrid(buildDrawContext(false));
                // 選択オブジェクトの処理 / Handle original selected objects
                if (cachedSelectionItems.length > 0 && isSelectionMode()) {
                    if (removeOriginalRadio && removeOriginalRadio.value) {
                        for (var sd = cachedSelectionItems.length - 1; sd >= 0; sd--) {
                            var itemToRemove = cachedSelectionItems[sd];
                            // ロック・非表示などで失敗しても他を続行 / Keep going even if one remove fails (locked/hidden, etc.)
                            safeExecute(function () { itemToRemove.remove(); });
                        }
                    } else if (makeOriginalGuidesRadio && makeOriginalGuidesRadio.value) {
                        // 元オブジェクトをガイド化 / Convert originals to guides
                        try {
                            doc.selection = null;
                            for (var sg = 0; sg < cachedSelectionItems.length; sg++) {
                                cachedSelectionItems[sg].selected = true;
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
                        for (var cc = 0; cc < drawnCellItems.length; cc++) {
                            drawnCellItems[cc].selected = true;
                        }
                        act_ShowCenter();
                        doc.selection = null;
                    }
                } catch (e) {
                    $.writeln("[GenerateGuidesGrid] show center error: " + e);
                }
            });
        } else {
            // Cancel: rollback preview changes
            previewMgr.rollback();
            // Cleanup preview layer just in case (fallback)
            safeRemoveLayerByName(doc, PREVIEW_LAYER_NAME);
            // プレビューで外れた選択を元に戻す / Restore selection lost during preview
            safeRestoreSelection();
        }
    }

    main();

})();