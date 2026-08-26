#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボードまたは選択オブジェクトの外接矩形を、指定した行数・列数に分割してグリッド用のガイドを生成します。
アートボードのエッジ、セルの長方形化、中心点の表示、プリセットの読み込み／書き出しにも対応します。

詳細は README を参照してください。

### Overview

Divides the artboard, or the bounding box of the selection, into the specified rows and columns and generates grid guides.
It can also draw the artboard edges, draw the cells as rectangles, mark center points, and import/export presets.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GenerateGuidesGrid";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-27";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GenerateGuidesGrid.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GenerateGuidesGrid.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n7adc7290b607"; /* 紹介記事 / article URL */

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
    var RECT_TOLERANCE = 0.5;                    /* アートボード内外判定の許容値（pt）/ Tolerance for the inside-artboard test (pt) */

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var ROW_SPACING    = 8;                  /* 行内の要素間隔 / gap inside a row */

    /**
     * ウィンドウの共通設定
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
     * パネルの共通設定
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
     * グループの共通設定（row は縦中央、column は左揃え）
     * @param {Group} group - 対象グループ
     * @param {string} [orientation] - "row" または "column"（省略時は "column"）
     * @param {number} [spacing] - 要素間隔（省略時は ROW_SPACING）
     * @returns {void}
     */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : ROW_SPACING;
    }

    /**
     * 行グループの共通設定（ボタン列・入力行など）
     * @param {Group} group - 対象グループ
     * @param {string} [alignment] - グループ自身の整列（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        /* alignment と alignChildren は対で指定する（片方だけだと天地がずれ、中のボタンが横に伸びる）
           Set both: alone, either one lets the row drift vertically or stretch its buttons */
        group.alignment = [alignment || "left", "center"];
        group.alignChildren = ["left", "center"];
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================
    function getUiLanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getUiLanguage();

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
            artboardEdge: { ja: "アートボードのエッジ", en: "Artboard Edges" },
            clearGuides: { ja: "既存ガイドを削除", en: "Clear Existing Guides" }
        },
        /* フィールド見出し（コロンは labelWithColon で付与）/ Field labels (colon added by labelWithColon) */
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
            guideExtension: { ja: "アートボード／オブジェクトの外側へガイドを伸ばす距離。", en: "Distance to mergeInto guides beyond the artboard/object." },
            artboardEdge: { ja: "アートボードの上下左右4辺にガイドを引きます。対象が選択オブジェクトのときは使えません。", en: "Draw guides on the four edges of the artboard. Unavailable when the target is the selected objects." },
            clearGuides: { ja: "描画前に専用レイヤー（grid_guides）内の既存ガイドを削除します。「すべてのアートボード」以外は、アクティブなアートボード上のガイドだけが対象です。", en: "Remove existing guides in the dedicated layer (grid_guides) before drawing. Unless all artboards are targeted, only guides on the active artboard are removed." },
            toGuide: { ja: "元の選択オブジェクトをガイドに変換します。", en: "Convert the original selected objects into guides." },
            outline: { ja: "アウトライン表示とプレビュー表示を切り替えます。", en: "Toggle between Outline and Preview view." },
            exportPreset: { ja: "現在の設定をプリセットファイルに書き出します。", en: "Export the current settings to a preset file." }
        },
        /* ボタン / Buttons */
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            apply: { ja: "適用", en: "Apply" },
            ok: { ja: "OK", en: "OK" },
            outline: { ja: "アウトライン表示", en: "Outline" },
            preview: { ja: "プレビュー表示", en: "Preview" },
            exportPreset: { ja: "書き出し", en: "Export" }
        }
    };

    /* ラベル取得（ドット区切りキー、{slash} は "/" に置換）/ Get label (dotted key, {slash} replaced with "/") */
    function getLabel(key) {
        var keyParts = key.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            if (labelNode && labelNode[keyParts[i]] !== undefined) {
                labelNode = labelNode[keyParts[i]];
            } else {
                return key;
            }
        }
        var labelValue = (labelNode && labelNode[currentLanguage] !== undefined) ? labelNode[currentLanguage] : key;
        return String(labelValue).replace(/\{slash\}/g, "/");
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelWithColon(key) {
        return getLabel(key) + (currentLanguage === "ja" ? "：" : ":");
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
    var UNIT_DECIMALS = 2; /* 換算時に残す小数桁数 / decimals kept when converting */

    /**
     * 指定桁数で丸める
     * @param {number} value - 対象の値
     * @param {number} decimals - 残す小数桁数
     * @returns {number} 丸めた値
     */
    function roundTo(value, decimals) {
        var digitScale = Math.pow(10, decimals);
        return Math.round(Number(value) * digitScale) / digitScale;
    }

    /**
     * pt → 現在単位
     * @param {number} ptValue - pt 値
     * @returns {number} 現在単位の値
     */
    function ptToUnit(ptValue) {
        return roundTo(Number(ptValue) / unitFactor, UNIT_DECIMALS);
    }

    /**
     * 現在単位 → pt
     * @param {number|string} unitValue - 現在単位の値
     * @returns {number} pt 値
     */
    function unitToPt(unitValue) {
        return roundTo(Number(unitValue) * unitFactor, UNIT_DECIMALS);
    }

    /**
     * 入力欄のテキストを数値として読む（空欄・不正値は fallback）
     * @param {string} inputText - 入力文字列
     * @param {number} [fallbackValue] - 数値にならないときの値（省略時は 0）
     * @returns {number} 数値
     */
    function toNumber(inputText, fallbackValue) {
        var parsedValue = parseFloat(inputText);
        if (isNaN(parsedValue)) return (typeof fallbackValue === "number") ? fallbackValue : 0;
        return parsedValue;
    }

    /**
     * 入力欄のテキストを整数として読む（空欄・不正値は fallback）
     * @param {string} inputText - 入力文字列
     * @param {number} [fallbackValue] - 整数にならないときの値（省略時は 0）
     * @returns {number} 整数
     */
    function toInteger(inputText, fallbackValue) {
        var parsedValue = parseInt(inputText, 10);
        if (isNaN(parsedValue)) return (typeof fallbackValue === "number") ? fallbackValue : 0;
        return parsedValue;
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
            columnGutter: 0,
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
            columnGutter: 0,
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
            columnGutter: 0,
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
            columnGutter: 0,
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
            columnGutter: 50,
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
            columnGutter: 30,
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
            columnGutter: 20,
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
            columnGutter: 20,
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
            columnGutter: 0,
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
            columnGutter: 0,
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
            columnGutter: 0,
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
    // func が false を返した手順はドキュメントを変更していないので数えない（余分な undo でユーザーの作業を巻き戻さない）
    // A step returning false changed nothing, so it is not counted (a surplus undo would roll back the user's own work)
    PreviewManager.prototype.addStep = function (step) {
        try {
            if (step() !== false) {
                this.undoDepth++;
            }
            app.redraw();
        } catch (e) {
            $.writeln("[GenerateGuidesGrid] preview step error: " + e);
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

    /**
     * 外接矩形（geometricBounds）を持つオブジェクトか
     * @param {object} item - 判定するオブジェクト
     * @returns {boolean} 座標を取得できるなら true
     */
    function hasGeometricBounds(item) {
        try {
            return !!(item && item.geometricBounds && item.geometricBounds.length === 4);
        } catch (e) {
            return false;
        }
    }

    /**
     * オブジェクトの中心が矩形の中にあるか（ガイドは矩形の外へ伸びるので中心で判定）
     * @param {PageItem} item - 判定するオブジェクト
     * @param {Array} boundsRect - 矩形 [左, 上, 右, 下]
     * @returns {boolean} 中心が矩形内なら true
     */
    function isCenterInsideRect(item, boundsRect) {
        if (!hasGeometricBounds(item)) return false;
        var itemBounds = item.geometricBounds;
        var centerX = (itemBounds[0] + itemBounds[2]) / 2;
        var centerY = (itemBounds[1] + itemBounds[3]) / 2;
        return (centerX >= boundsRect[0] - RECT_TOLERANCE && centerX <= boundsRect[2] + RECT_TOLERANCE &&
            centerY <= boundsRect[1] + RECT_TOLERANCE && centerY >= boundsRect[3] - RECT_TOLERANCE);
    }

    /**
     * 関数を実行し、例外を握りつぶす（try/catch の重複を集約）
     * @param {function} action - 実行する処理
     * @returns {void}
     */
    function safeExecute(action) {
        try {
            action();
        } catch (e) {
            $.writeln("[GenerateGuidesGrid] safeExecute error: " + e);
        }
    }

    /**
     * source の自前プロパティを target にコピーする
     * @param {object} target - コピー先
     * @param {object} source - コピー元
     * @returns {object} コピー先（target）
     */
    function mergeInto(target, source) {
        for (var key in source) {
            if (source.hasOwnProperty(key)) target[key] = source[key];
        }
        return target;
    }

    /**
     * レイヤーのロックを安全に解除する
     * @param {Layer} layer - 対象レイヤー
     * @returns {void}
     */
    function safeUnlockLayer(layer) {
        safeExecute(function () {
            if (layer && layer.locked) layer.locked = false;
        });
    }

    /**
     * 指定名のレイヤーを安全に削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {void}
     */
    function safeRemoveLayerByName(doc, layerName) {
        safeExecute(function () {
            var layer = doc.layers.getByName(layerName);
            if (layer) layer.remove();
        });
    }

    /**
     * 指定名のレイヤーを取得する。なければ作成する
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 取得または作成したレイヤー
     */
    function getOrCreateLayer(doc, layerName) {
        var layer;
        try {
            layer = doc.layers.getByName(layerName);
        } catch (e) {
            layer = doc.layers.add();
            layer.name = layerName;
        }
        return layer;
    }

    /**
     * ガイド線を1本追加する（塗り・線なしのパスをガイド化してレイヤー先頭へ）
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} layer - 追加先レイヤー
     * @param {Array} startPoint - 始点 [x, y]
     * @param {Array} endPoint - 終点 [x, y]
     * @returns {PathItem} 追加したガイド
     */
    function addGuideLine(doc, layer, startPoint, endPoint) {
        var guideLine = doc.pathItems.add();
        guideLine.setEntirePath([startPoint, endPoint]);
        guideLine.stroked = false;
        guideLine.filled = false;
        guideLine.guides = true;
        guideLine.move(layer, ElementPlacement.PLACEATBEGINNING);
        return guideLine;
    }

    /**
     * 角丸ライブエフェクトのXMLを作る
     * @param {number} radius - 角丸半径（pt）
     * @returns {string} ライブエフェクトのXML
     */
    function roundCornersEffectXML(radius) {
        var effectXml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius #value# "/></LiveEffect>';
        return effectXml.replace('#value#', radius);
    }

    /**
     * 黒色を作る（CMYK／RGB対応）
     * @param {Document} doc - 対象ドキュメント
     * @returns {CMYKColor|RGBColor} 黒のカラーオブジェクト
     */
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

    /**
     * 描画できるコンテキストか（行数・列数が正か）
     * @param {object} drawContext - 描画コンテキスト
     * @returns {boolean} 描画できるなら true
     */
    function canDrawGrid(drawContext) {
        return !(isNaN(drawContext.columnCount) || drawContext.columnCount <= 0 || isNaN(drawContext.rowCount) || drawContext.rowCount <= 0);
    }

    // グリッド（ガイド＋セル長方形）を描画し、作成したセル長方形の配列を返す / Draw the grid; return created cell rects
    function drawGrid(drawContext) {
        var doc = drawContext.doc;
        var isPreview = drawContext.isPreview;
        var columnCount = drawContext.columnCount;
        var rowCount = drawContext.rowCount;
        if (!canDrawGrid(drawContext)) return [];

        var guideExtension = drawContext.guideExtension;
        var marginTop = drawContext.marginTop, marginBottom = drawContext.marginBottom, marginLeft = drawContext.marginLeft, marginRight = drawContext.marginRight;
        var rowGutter = drawContext.rowGutter, columnGutter = drawContext.columnGutter;
        var drawCells = drawContext.drawCells, drawGridGuides = drawContext.drawGridGuides, cornerRadius = drawContext.cornerRadius;
        var splitCells = drawContext.splitCells;
        // 選択オブジェクトが対象のときはアートボードの辺を引かない / No artboard edges when the target is the selection
        var drawArtboardEdge = drawContext.drawArtboardEdge && !drawContext.selBounds;
        var cellOpacity = (typeof drawContext.cellOpacity === "number") ? drawContext.cellOpacity : CELL_OPACITY;
        var createdCells = [];

        // プレビューは1枚に全部描く。確定時はガイド用レイヤーはガイド系を描くときだけ作る
        // Preview draws everything on one layer; on commit, create the guide layer only when guides are drawn
        var gridLayerName = isPreview ? PREVIEW_LAYER_NAME : GUIDE_LAYER_NAME;
        var needGridLayer = isPreview ? (drawGridGuides || splitCells || drawCells || drawArtboardEdge) : (drawGridGuides || splitCells || drawArtboardEdge);
        var gridLayer = needGridLayer ? getOrCreateLayer(doc, gridLayerName) : null;
        safeUnlockLayer(gridLayer);

        var cellLayer = gridLayer; // プレビューはセルも gridLayer に描く / preview: cells live on gridLayer
        if (!isPreview && drawCells) {
            cellLayer = getOrCreateLayer(doc, CELL_LAYER_NAME);
            safeUnlockLayer(cellLayer);
        }

        var targetRects = [];
        if (drawContext.selBounds) {
            targetRects.push(drawContext.selBounds);
        } else {
            for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
                if (!drawContext.allBoards && artboardIndex !== doc.artboards.getActiveArtboardIndex()) continue;
                targetRects.push(doc.artboards[artboardIndex].artboardRect);
            }
        }

        for (var targetIndex = 0; targetIndex < targetRects.length; targetIndex++) {
            var targetRect = targetRects[targetIndex];
            var targetLeft = targetRect[0],
                targetTop = targetRect[1],
                targetRight = targetRect[2],
                targetBottom = targetRect[3];
            var contentLeft = targetLeft + marginLeft;
            var contentRight = targetRight - marginRight;
            var contentTop = targetTop - marginTop;
            var contentBottom = targetBottom + marginBottom;

            var extendedLeft = targetLeft - guideExtension;
            var extendedRight = targetRight + guideExtension;
            var extendedTop = targetTop + guideExtension;
            var extendedBottom = targetBottom - guideExtension;

            // アートボードの上下左右4辺（マージンに関係なく引く）/ The artboard's four edges (independent of the margins)
            if (drawArtboardEdge && gridLayer) {
                addGuideLine(doc, gridLayer, [extendedLeft, targetTop], [extendedRight, targetTop]);
                addGuideLine(doc, gridLayer, [extendedLeft, targetBottom], [extendedRight, targetBottom]);
                addGuideLine(doc, gridLayer, [targetLeft, extendedTop], [targetLeft, extendedBottom]);
                addGuideLine(doc, gridLayer, [targetRight, extendedTop], [targetRight, extendedBottom]);
            }

            var usableWidth = contentRight - contentLeft;
            var usableHeight = contentTop - contentBottom;
            var totalColumnGutter = (columnCount - 1) * columnGutter;
            var totalRowGutter = (rowCount - 1) * rowGutter;
            // マージン・ガターが過大でセル幅/高さが0以下になる場合はこの対象をスキップ
            // Skip this target if margins/gutters are too large (cell width/height would be <= 0)
            if (usableWidth - totalColumnGutter <= 0 || usableHeight - totalRowGutter <= 0) continue;
            var cellWidth = (usableWidth - totalColumnGutter) / columnCount;
            var cellHeight = (usableHeight - totalRowGutter) / rowCount;

            if (drawGridGuides) {
                if (columnCount === 1 && rowCount === 1) {
                    // 四辺（マージン適用後の有効領域）をガイド化 / Four edges of the usable area
                    addGuideLine(doc, gridLayer, [extendedLeft, contentTop], [extendedRight, contentTop]);
                    addGuideLine(doc, gridLayer, [extendedLeft, contentBottom], [extendedRight, contentBottom]);
                    addGuideLine(doc, gridLayer, [contentLeft, extendedTop], [contentLeft, extendedBottom]);
                    addGuideLine(doc, gridLayer, [contentRight, extendedTop], [contentRight, extendedBottom]);
                } else {
                    // 通常ガイド描画（行・列）/ Normal grid guides (rows and columns)
                    // 最終行は contentBottom にちょうど着地するので、末尾に足すと重なる
                    // The last row lands exactly on contentBottom, so a trailing guide would overlap
                    var lineY = contentTop;
                    addGuideLine(doc, gridLayer, [extendedLeft, lineY], [extendedRight, lineY]);
                    for (var j = 0; j < rowCount; j++) {
                        lineY -= cellHeight;
                        addGuideLine(doc, gridLayer, [extendedLeft, lineY], [extendedRight, lineY]);
                        // ガター0のときは同じ位置に重なるので引かない / Gutter 0 would draw on the same line
                        if (j < rowCount - 1 && rowGutter > 0) {
                            lineY -= rowGutter;
                            addGuideLine(doc, gridLayer, [extendedLeft, lineY], [extendedRight, lineY]);
                        }
                    }

                    var lineX = contentLeft;
                    addGuideLine(doc, gridLayer, [lineX, extendedTop], [lineX, extendedBottom]);
                    for (var k = 0; k < columnCount; k++) {
                        lineX += cellWidth;
                        addGuideLine(doc, gridLayer, [lineX, extendedTop], [lineX, extendedBottom]);
                        // ガター0のときは同じ位置に重なるので足さない / Gutter 0 would draw on the same line
                        if (k < columnCount - 1 && columnGutter > 0) {
                            lineX += columnGutter;
                            addGuideLine(doc, gridLayer, [lineX, extendedTop], [lineX, extendedBottom]);
                        }
                    }
                }
            }

            if (drawCells && cellLayer) {
                var cellOriginX = contentLeft;
                var cellOriginY = contentTop;
                for (var row = 0; row < rowCount; row++) {
                    var cellY = cellOriginY - (cellHeight + rowGutter) * row;
                    for (var column = 0; column < columnCount; column++) {
                        var cellX = cellOriginX + (cellWidth + columnGutter) * column;
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
                var splitOriginX = contentLeft;
                var splitOriginY = contentTop;
                for (var splitRow = 0; splitRow < rowCount; splitRow++) {
                    var splitCellTop = splitOriginY - (cellHeight + rowGutter) * splitRow;
                    var splitCellBottom = splitCellTop - cellHeight;
                    for (var splitColumn = 0; splitColumn < columnCount; splitColumn++) {
                        var splitCenterX = splitOriginX + (cellWidth + columnGutter) * splitColumn + cellWidth / 2;
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
        var previewManager = new PreviewManager();

        // 確定描画したセル長方形（中心点表示の対象）/ Cell rects from final draw (targets for center-point display)
        var drawnCellItems = [];

        // 選択オブジェクトの参照と外接矩形をキャッシュ / Cache selection refs and bounds before dialog
        var cachedSelectionItems = [];
        var cachedSelectionBounds = (function () {
            var selection = doc.selection;
            // 文字選択中（TextRange）は座標を持たないので対象外 / A TextRange has no bounds, so it is not a target
            if (!selection || selection.typename === "TextRange" || selection.length === 0) return null;
            for (var i = 0; i < selection.length; i++) {
                if (hasGeometricBounds(selection[i])) cachedSelectionItems.push(selection[i]);
            }
            if (cachedSelectionItems.length === 0) return null;
            var firstItemBounds = cachedSelectionItems[0].geometricBounds;
            var selectionLeft = firstItemBounds[0], selectionTop = firstItemBounds[1], selectionRight = firstItemBounds[2], selectionBottom = firstItemBounds[3];
            for (var j = 1; j < cachedSelectionItems.length; j++) {
                var itemBounds = cachedSelectionItems[j].geometricBounds;
                if (itemBounds[0] < selectionLeft) selectionLeft = itemBounds[0];
                if (itemBounds[1] > selectionTop) selectionTop = itemBounds[1];
                if (itemBounds[2] > selectionRight) selectionRight = itemBounds[2];
                if (itemBounds[3] < selectionBottom) selectionBottom = itemBounds[3];
            }
            return [selectionLeft, selectionTop, selectionRight, selectionBottom];
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
        setupWindow(dialog);

        // プリセット選択＋書き出し / Preset selection and export
        var presetGroup = dialog.add("group");
        setupGroup(presetGroup, "row");
        presetGroup.alignment = ["center", "top"]; // 左右中央 / horizontally centered
        presetGroup.margins = [0, 0, 0, 10]; // 下に少し余白 / add some bottom margin

        presetGroup.add("statictext", undefined, labelWithColon("field.preset"));
        var presetDropdown = presetGroup.add("dropdownlist", undefined, []);
        presetDropdown.selection = 0;
        var btnExportPreset = presetGroup.add("button", undefined, getLabel("button.exportPreset"));
        btnExportPreset.alignment = ["left", "center"]; // 横に伸ばさず天地中央 / natural width, vertically centered
        btnExportPreset.helpTip = getLabel("tooltip.exportPreset");

        // プリセット書き出し / Export the current settings as a preset
        btnExportPreset.onClick = function () {
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
                columns: toInteger(columnCountInput.text, 1),
                rows: toInteger(rowCountInput.text, 1),
                guideExtension: unitToPt(toNumber(extensionInput.text, 0)),
                marginTop: unitToPt(toNumber(marginTopInput.text, 0)),
                marginBottom: unitToPt(toNumber(marginBottomInput.text, 0)),
                marginLeft: unitToPt(toNumber(marginLeftInput.text, 0)),
                marginRight: unitToPt(toNumber(marginRightInput.text, 0)),
                rowGutter: unitToPt(toNumber(rowGutterInput.text, 0)),
                columnGutter: unitToPt(toNumber(columnGutterInput.text, 0)),
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
                'columnGutter: ' + currentPreset.columnGutter + ', ' +
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
        setupGroup(targetRow, "row", COLUMN_SPACING);
        targetRow.alignChildren = ["left", "fill"];

        // 対象パネル：対象の種類（縦並び）/ Target panel: target type (vertical)
        var targetPanel = targetRow.add("panel", undefined, getLabel("panel.target"));
        setupPanel(targetPanel);
        var targetTypeGroup = targetPanel.add("group");
        setupGroup(targetTypeGroup, "column");
        var targetSelectionRadio = targetTypeGroup.add("radiobutton", undefined, getLabel("target.selection"));
        var targetArtboardRadio = targetTypeGroup.add("radiobutton", undefined, getLabel("target.artboard"));
        var targetAllArtboardsRadio = targetTypeGroup.add("radiobutton", undefined, getLabel("target.allArtboards"));

        // 元のオブジェクトパネル（対象パネルの外・横並び、選択オブジェクト時のみ有効）/ Original-object panel (outside target panel, selection mode only)
        var originalObjectPanel = targetRow.add("panel", undefined, getLabel("panel.originalObject"));
        setupPanel(originalObjectPanel);
        var originalObjectGroup = originalObjectPanel.add("group");
        setupGroup(originalObjectGroup, "column");
        var removeOriginalRadio = originalObjectGroup.add("radiobutton", undefined, getLabel("radio.remove"));
        var keepOriginalRadio = originalObjectGroup.add("radiobutton", undefined, getLabel("radio.keep"));
        var makeOriginalGuidesRadio = originalObjectGroup.add("radiobutton", undefined, getLabel("radio.toGuide"));
        makeOriginalGuidesRadio.helpTip = getLabel("tooltip.toGuide");
        keepOriginalRadio.value = true; // デフォルト / default

        // 選択がなければ「選択オブジェクト」を無効化 / Disable selection option when nothing is selected
        if (!cachedSelectionBounds) {
            targetSelectionRadio.enabled = false;
        }
        // アートボードが1つだけなら「すべてのアートボード」を無効化 / Disable "All Artboards" when only one artboard
        if (!hasMultipleArtboards) {
            targetAllArtboardsRadio.enabled = false;
        }
        // 初期選択 / Initial target radio
        targetSelectionRadio.value = (targetMode === "selection");
        targetArtboardRadio.value = (targetMode === "artboard");
        targetAllArtboardsRadio.value = (targetMode === "allArtboards");

        // 対象モード変更 / Target mode change
        function onTargetChange() {
            if (targetSelectionRadio.value) {
                targetMode = "selection";
            } else if (targetAllArtboardsRadio.value) {
                targetMode = "allArtboards";
            } else {
                targetMode = "artboard";
            }
            updateTargetMode();
            safeUpdatePreview();
        }
        targetSelectionRadio.onClick = targetArtboardRadio.onClick = targetAllArtboardsRadio.onClick = onTargetChange;

        // 元オブジェクト処理の切替でプレビュー更新 / Update preview on original-object radio change
        removeOriginalRadio.onClick = keepOriginalRadio.onClick = makeOriginalGuidesRadio.onClick = function () {
            safeUpdatePreview();
        };

        // 対象モードに応じた表示制御 / Enable/disable controls by target mode
        function updateTargetMode() {
            var isSelectionTarget = isSelectionMode();
            originalObjectPanel.enabled = isSelectionTarget;
            // 選択オブジェクトにはアートボードの辺がないのでディム / No artboard edges when the target is the selection
            artboardEdgeCheckbox.enabled = !isSelectionTarget;
        }

        // グリッド設定グループ / Grid settings group
        var gridSettingRow = dialog.add("group");
        setupGroup(gridSettingRow, "row", COLUMN_SPACING);
        gridSettingRow.alignChildren = ["left", "top"];
        var gridLabelWidth = (currentLanguage === "ja") ? 40 : 50; // unify Number/Gutter label width and right-align

        // 行設定パネル / Row settings panel
        var rowSettingPanel = gridSettingRow.add("panel", undefined, getLabel("panel.row"));
        setupPanel(rowSettingPanel);

        var rowCountGroup = rowSettingPanel.add("group");
        setupGroup(rowCountGroup, "row");
        var rowCountLabel = rowCountGroup.add("statictext", undefined, labelWithColon("field.rowCount"));
        rowCountLabel.justification = "right";
        rowCountLabel.minimumSize.width = gridLabelWidth;
        rowCountLabel.maximumSize.width = gridLabelWidth;
        var rowCountInput = rowCountGroup.add("edittext", undefined, "2");
        rowCountInput.characters = 3;

        var rowGutterGroup = rowSettingPanel.add("group");
        setupGroup(rowGutterGroup, "row");
        var rowGutterLabel = rowGutterGroup.add("statictext", undefined, labelWithColon("field.rowGutter"));
        rowGutterLabel.justification = "right";
        rowGutterLabel.minimumSize.width = gridLabelWidth;
        rowGutterLabel.maximumSize.width = gridLabelWidth;
        var rowGutterInput = rowGutterGroup.add("edittext", undefined, "0");
        rowGutterInput.characters = 3;
        rowGutterGroup.add("statictext", undefined, unitLabel);

        // 列設定パネル / Column settings panel
        var columnSettingPanel = gridSettingRow.add("panel", undefined, getLabel("panel.column"));
        setupPanel(columnSettingPanel);

        var columnCountGroup = columnSettingPanel.add("group");
        setupGroup(columnCountGroup, "row");
        var columnCountLabel = columnCountGroup.add("statictext", undefined, labelWithColon("field.columnCount"));
        columnCountLabel.justification = "right";
        columnCountLabel.minimumSize.width = gridLabelWidth;
        columnCountLabel.maximumSize.width = gridLabelWidth;
        var columnCountInput = columnCountGroup.add("edittext", undefined, "2");
        columnCountInput.characters = 3;

        var columnGutterGroup = columnSettingPanel.add("group");
        setupGroup(columnGutterGroup, "row");
        var columnGutterLabel = columnGutterGroup.add("statictext", undefined, labelWithColon("field.columnGutter"));
        columnGutterLabel.justification = "right";
        columnGutterLabel.minimumSize.width = gridLabelWidth;
        columnGutterLabel.maximumSize.width = gridLabelWidth;
        var columnGutterInput = columnGutterGroup.add("edittext", undefined, "0");
        columnGutterInput.characters = 3;
        columnGutterGroup.add("statictext", undefined, unitLabel);

        // 行間に連動（列パネル下部）/ Link to row gutter (under column panel)
        var linkGutterCheckbox = columnSettingPanel.add("checkbox", undefined, getLabel("checkbox.linkGutter"));
        linkGutterCheckbox.helpTip = getLabel("tooltip.linkGutter");
        linkGutterCheckbox.value = true;

        // 連動ON時は列間を行間と同じ値にする / When linked, sync col gutter to row gutter
        function syncGutterLink() {
            if (linkGutterCheckbox.value) {
                columnGutterInput.text = rowGutterInput.text;
                columnGutterGroup.enabled = false;
            } else {
                var columnCount = parseInt(columnCountInput.text, 10);
                columnGutterGroup.enabled = (columnCount > 1);
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
            var cellGroup = parentRow.add("group");
            cellGroup.orientation = "row";
            cellGroup.alignment = ["center", "center"];
            cellGroup.minimumSize.width = MARGIN_CELL_WIDTH;
            cellGroup.add("statictext", undefined, labelWithColon(labelKey));
            var marginInput = cellGroup.add("edittext", undefined, "0");
            marginInput.characters = 3;
            return { group: cellGroup, input: marginInput };
        }
        // 位置合わせ用の空セル / Empty cell for alignment
        function addMarginSpacer(parentRow) {
            var spacerGroup = parentRow.add("group");
            spacerGroup.minimumSize.width = MARGIN_CELL_WIDTH;
        }

        // 1行目：［空］［上］［空］/ Row 1: [empty][top][empty]
        var marginTopRow = marginPanel.add("group");
        marginTopRow.orientation = "row";
        marginTopRow.alignment = ["center", "top"];
        addMarginSpacer(marginTopRow);
        var marginTopCell = addMarginCell(marginTopRow, "field.top");
        var marginTopInput = marginTopCell.input;
        addMarginSpacer(marginTopRow);

        // 2行目：［左］［連動］［右］/ Row 2: [left][link][right]
        var marginMiddleRow = marginPanel.add("group");
        marginMiddleRow.orientation = "row";
        marginMiddleRow.alignment = ["center", "top"];
        var marginLeftCell = addMarginCell(marginMiddleRow, "field.left");
        var marginLeftGroup = marginLeftCell.group;
        var marginLeftInput = marginLeftCell.input;
        var linkMarginGroup = marginMiddleRow.add("group");
        linkMarginGroup.orientation = "row";
        linkMarginGroup.alignment = ["center", "center"];
        linkMarginGroup.minimumSize.width = MARGIN_CELL_WIDTH;
        var linkMarginCheckbox = linkMarginGroup.add("checkbox", undefined, getLabel("checkbox.linkMargin"));
        linkMarginCheckbox.helpTip = getLabel("tooltip.linkMargin");
        linkMarginCheckbox.value = true; // デフォルトでON / on by default
        var marginRightCell = addMarginCell(marginMiddleRow, "field.right");
        var marginRightGroup = marginRightCell.group;
        var marginRightInput = marginRightCell.input;

        // 3行目：［空］［下］［空］/ Row 3: [empty][bottom][empty]
        var marginBottomRow = marginPanel.add("group");
        marginBottomRow.orientation = "row";
        marginBottomRow.alignment = ["center", "top"];
        addMarginSpacer(marginBottomRow);
        var marginBottomCell = addMarginCell(marginBottomRow, "field.bottom");
        var marginBottomGroup = marginBottomCell.group;
        var marginBottomInput = marginBottomCell.input;
        addMarginSpacer(marginBottomRow);

        // セルパネル＋ガイドパネルを横並び（左：セル、右：ガイド）/ Cell panel + Guides panel side by side (left: cell, right: guides)
        var cellGuideRow = dialog.add("group");
        setupGroup(cellGuideRow, "row", COLUMN_SPACING);
        cellGuideRow.alignChildren = ["left", "fill"];

        // セルパネル（ガイドパネルの左）/ Cell panel (left of Guides panel)
        var cellPanel = cellGuideRow.add("panel", undefined, getLabel("panel.options"));
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
        var roundCornerInput = roundCornerGroup.add("edittext", undefined, "3");
        roundCornerInput.characters = 2;
        roundCornerGroup.add("statictext", undefined, unitLabel);

        // 不透明度スライダー（0-100、ラベルの次の行にスライダー）/ Opacity slider (0-100, slider on the line below the label)
        var opacityGroup = cellPanel.add("group");
        setupGroup(opacityGroup, "column", ROW_SPACING);
        opacityGroup.add("statictext", undefined, labelWithColon("field.opacity"));
        var opacitySlider = opacityGroup.add("slider", undefined, CELL_OPACITY, 0, 100);
        opacitySlider.helpTip = getLabel("tooltip.opacity");
        opacitySlider.alignment = ["fill", "center"];
        opacitySlider.onChanging = function () {
            safeUpdatePreview();
        };

        // ガイドパネル / Guides panel
        var guidesPanel = cellGuideRow.add("panel", undefined, getLabel("panel.guides"));
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

        // アートボードのエッジ（上下左右4辺）/ Artboard edges (all four sides)
        var artboardEdgeCheckbox = guidesPanel.add("checkbox", undefined, getLabel("checkbox.artboardEdge"));
        artboardEdgeCheckbox.helpTip = getLabel("tooltip.artboardEdge");
        artboardEdgeCheckbox.value = false;
        artboardEdgeCheckbox.onClick = function () {
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
        clearGuidesCheckbox.value = false;
        clearGuidesCheckbox.onClick = function () {
            safeUpdatePreview();
        };

        // 数値フィールドを矢印キーで増減（Shift で10の倍数にスナップ）/ Adjust a numeric field with arrow keys (Shift snaps to multiples of 10)
        function changeValueByArrowKey(editText) {
            editText.addEventListener("keydown", function (event) {
                var fieldValue = Number(editText.text);
                if (isNaN(fieldValue)) return;

                var keyboard = ScriptUI.environment.keyboardState;

                if (event.keyName == "Up" || event.keyName == "Down") {
                    var isUp = event.keyName == "Up";
                    var delta = 1;

                    if (keyboard.shiftKey) {
                        // Shiftキー押下時は10の倍数にスナップ
                        fieldValue = Math.floor(fieldValue / 10) * 10;
                        delta = 10;
                    }

                    fieldValue += isUp ? delta : -delta;
                    if (fieldValue < 0) fieldValue = 0; // 必要なら下限チェック

                    event.preventDefault();
                    editText.text = fieldValue;
                    // 連動・ガター・列間同期をまとめて反映 / Apply linked margin / gutter enable / col gutter sync together
                    if (editText === marginTopInput && linkMarginCheckbox.value) {
                        syncLinkedMargins();
                    }
                    if (editText === columnCountInput || editText === rowCountInput) {
                        updateGutterEnabled();
                    }
                    if (editText === rowGutterInput && linkGutterCheckbox.value) {
                        columnGutterInput.text = rowGutterInput.text;
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
                previewManager.rollback();
            } catch (e) {
                $.writeln("[GenerateGuidesGrid] preview rollback error: " + e);
            }

            // 「既存ガイドを削除」ONなら、この時点で削除して見た目に反映する
            // キャンセル時は rollback の app.undo() で元に戻る / Cancel restores them via rollback
            if (clearGuidesCheckbox.value) {
                previewManager.addStep(function () {
                    return clearExistingGuides(); // 何も消さなければ false / false when nothing was removed
                });
            }

            // Draw preview as one undoable step
            previewManager.addStep(function () {
                var drawContext = buildDrawContext(true);
                if (!canDrawGrid(drawContext)) return false; // 何も描かない＝undo対象なし / nothing drawn, nothing to undo
                drawGrid(drawContext); // プレビュー描画 / draw as preview
                return true;
            });

            // 選択オブジェクトの表示/非表示をプレビュー / Preview hide/show of selected objects
            if (isSelectionMode() && cachedSelectionItems.length > 0) {
                var shouldHide = (removeOriginalRadio && removeOriginalRadio.value);
                if (shouldHide) {
                    previewManager.addStep(function () {
                        for (var i = 0; i < cachedSelectionItems.length; i++) {
                            cachedSelectionItems[i].hidden = true;
                        }
                        return true;
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
        changeValueByArrowKey(marginTopInput);
        changeValueByArrowKey(marginBottomInput);
        changeValueByArrowKey(marginLeftInput);
        changeValueByArrowKey(marginRightInput);
        changeValueByArrowKey(rowGutterInput);
        changeValueByArrowKey(columnGutterInput);

        // 入力中の変更もリアルタイム反映 / Attach onChanging for live preview
        attachLivePreview(columnCountInput);
        attachLivePreview(rowCountInput);
        attachLivePreview(extensionInput);
        // 上マージン変更時に連動ONなら左右下も同期 / Sync margins when top changes (if linked)
        marginTopInput.onChanging = function () {
            if (linkMarginCheckbox.value) {
                marginBottomInput.text = marginTopInput.text;
                marginLeftInput.text = marginTopInput.text;
                marginRightInput.text = marginTopInput.text;
            }
            safeUpdatePreview();
        };
        attachLivePreview(marginBottomInput);
        attachLivePreview(marginLeftInput);
        attachLivePreview(marginRightInput);
        // 行間変更時に連動チェックONなら列間も同期 / Sync col gutter when row gutter changes (if linked)
        rowGutterInput.onChanging = function () {
            if (linkGutterCheckbox.value) {
                columnGutterInput.text = rowGutterInput.text;
            }
            safeUpdatePreview();
        };
        attachLivePreview(columnGutterInput);

        // === ボタンエリア（3カラム：左アウトライン／中央スペーサー／右キャンセル・OK）/ Button area (3 columns: left outline / center spacer / right cancel+ok)
        var btnRowGroup = dialog.add("group");
        btnRowGroup.alignment = ["fill", "top"];
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["fill", "center"];
        btnRowGroup.margins = [0, 5, 0, 0]; // ボタンエリア上マージン +5 / extra top margin
        btnRowGroup.spacing = 0;

        // 左グループ（アウトラインボタン）/ Left group (Outline button)
        var btnLeftGroup = btnRowGroup.add("group");
        setupRow(btnLeftGroup, "left");
        var btnOutline = btnLeftGroup.add("button", undefined, getLabel("button.outline"));
        btnOutline.helpTip = getLabel("tooltip.outline");
        // アウトライン⇔プレビュー表示を切り替え、ラベルもトグル / Toggle Outline/Preview view and the button label
        btnOutline.onClick = function () {
            app.executeMenuCommand('preview');
            btnOutline.text = (btnOutline.text === getLabel("button.outline"))
                ? getLabel("button.preview")
                : getLabel("button.outline");
        };

        // スペーサー（横に伸びる空白）/ Spacer (horizontal stretch)
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;
        spacer.maximumSize.height = 0;

        // 右グループ（キャンセル・OKボタン）/ Right group (Cancel/OK buttons)
        var btnRightGroup = btnRowGroup.add("group");
        setupRow(btnRightGroup, "right", 10);
        btnRightGroup.alignChildren = ["right", "center"];
        var btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), {
            name: "cancel"
        });
        var btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), {
            name: "ok"
        });
        btnOK.alignment = ["right", "center"];

        // 表示用ラベルをローカライズ / Localize display label for dropdown
        function presetDisplayLabel(rawLabel) {
            // 日本語UIのときは「 / 」以降を隠す / In Japanese UI, hide text after " / "
            if (currentLanguage === "ja") return String(rawLabel).replace(/\s*\/.*$/, "");
            return rawLabel;
        }

        // プリセットをドロップダウンに追加 / Add presets to dropdown
        for (var i = 0; i < presets.length; i++) {
            var presetLabel = presetDisplayLabel(presets[i].label);
            presetDropdown.add("item", presetLabel);
        }
        presetDropdown.selection = 0;

        // プリセットの値を入力欄に反映する共通関数 / Common function to apply preset values
        // 旧キー（x/y/top/bottom/left/right）も後方互換で受け付ける / Accept legacy keys for backward compatibility
        function applyPreset(preset) {
            /* 値を取得（新キー→旧キー→既定値の順）/ Pick a value: new key → legacy key → default */
            function pickPresetValue(preferred, legacy, fallbackValue) {
                if (preferred !== undefined) return preferred;
                if (legacy !== undefined) return legacy;
                return fallbackValue;
            }
            // 行数・列数（換算なし）/ Counts (no unit conversion)
            columnCountInput.text = pickPresetValue(preset.columns, preset.x, 1);
            rowCountInput.text = pickPresetValue(preset.rows, preset.y, 1);
            // 長さ系は pt 基準なので現在単位へ換算 / Length values are stored in pt — convert to the current unit
            extensionInput.text = ptToUnit(pickPresetValue(preset.guideExtension, undefined, 0));
            var presetMarginTop = pickPresetValue(preset.marginTop, preset.top, 0);
            var presetMarginBottom = pickPresetValue(preset.marginBottom, preset.bottom, 0);
            var presetMarginLeft = pickPresetValue(preset.marginLeft, preset.left, 0);
            var presetMarginRight = pickPresetValue(preset.marginRight, preset.right, 0);
            marginTopInput.text = ptToUnit(presetMarginTop);
            marginBottomInput.text = ptToUnit(presetMarginBottom);
            marginLeftInput.text = ptToUnit(presetMarginLeft);
            marginRightInput.text = ptToUnit(presetMarginRight);
            rowGutterInput.text = ptToUnit(pickPresetValue(preset.rowGutter, undefined, 0));
            columnGutterInput.text = ptToUnit(pickPresetValue(preset.columnGutter, undefined, 0));
            // 上下左右が異なるプリセットは連動をOFF（連動が値を上書きして壊すのを防ぐ）
            // If margins differ, turn the link off so it won't overwrite the distinct values
            linkMarginCheckbox.value = (presetMarginTop === presetMarginBottom && presetMarginTop === presetMarginLeft && presetMarginTop === presetMarginRight);
            cellRectCheckbox.value = (typeof preset.drawCells !== "undefined") ? preset.drawCells : false;
            drawGuidesCheckbox.value = (typeof preset.drawGuides !== "undefined") ? preset.drawGuides : true;
            extensionGroup.enabled = drawGuidesCheckbox.value;
        }

        // プリセット選択時に入力値へ反映 / Apply preset values to inputs on selection
        presetDropdown.onChange = function () {
            applyPreset(presets[presetDropdown.selection.index]);
            updateGutterEnabled();
            updateTargetMode();
            updateCellOptionEnabled();
            syncLinkedMargins();
            safeUpdatePreview();
        };

        // 初期プリセットの値を入力欄に反映 / Apply initial preset values to input fields
        applyPreset(presets[0]);

        // 「連動」同期処理 / Sync for "Link" margin
        function syncLinkedMargins() {
            if (linkMarginCheckbox.value) {
                var topValue = marginTopInput.text;
                marginBottomInput.text = topValue;
                marginLeftInput.text = topValue;
                marginRightInput.text = topValue;
                marginBottomGroup.enabled = false;
                marginLeftGroup.enabled = false;
                marginRightGroup.enabled = false;
            } else {
                marginBottomGroup.enabled = true;
                marginLeftGroup.enabled = true;
                marginRightGroup.enabled = true;
            }
        }
        linkMarginCheckbox.onClick = function () {
            syncLinkedMargins();
            safeUpdatePreview();
        };

        // ガター有効無効切り替え / Enable/disable gutter fields
        function updateGutterEnabled() {
            var columnCount = parseInt(columnCountInput.text, 10);
            var rowCount = parseInt(rowCountInput.text, 10);
            rowGutterGroup.enabled = (rowCount > 1);
            // 行数・列数が両方2以上のときのみ連動を有効 / Enable link only when both >= 2
            var hasMultipleRowsAndColumns = (columnCount > 1 && rowCount > 1);
            linkGutterCheckbox.enabled = hasMultipleRowsAndColumns;
            if (linkGutterCheckbox.value && hasMultipleRowsAndColumns) {
                columnGutterInput.text = rowGutterInput.text;
                columnGutterGroup.enabled = false;
            } else {
                columnGutterGroup.enabled = (columnCount > 1);
            }
        }

        // 行数・列数変更時：ガター有効/無効を更新して即時プレビュー / On row/col change: refresh gutter enable + preview
        columnCountInput.onChanging = rowCountInput.onChanging = function () {
            updateGutterEnabled();
            safeUpdatePreview();
        };

        // 「ガイドを引く」切替：伸張・クリアをディム制御してプレビュー / Toggle draw-guides: dim extension/clear, then preview
        drawGuidesCheckbox.onClick = function () {
            var guidesEnabled = drawGuidesCheckbox.value;
            extensionGroup.enabled = guidesEnabled;
            extensionInput.enabled = guidesEnabled && extensionCheckbox.value;
            clearGuidesCheckbox.enabled = guidesEnabled; // ガイドOFFならクリアをディム / dim clear when guides off
            safeUpdatePreview();
        };

        // 長方形化の有無でセル系オプション（中心点・角丸・不透明度）をディム制御 / Cell options dim unless rectangles are drawn
        // 角丸の数値欄は「長方形化ON かつ 角丸ON」のときだけ有効 / The round-corner field is enabled only when both rectangles and round corners are on
        function updateCellOptionEnabled() {
            var cellsEnabled = cellRectCheckbox.value;
            centerPointCheckbox.enabled = cellsEnabled;
            roundCornerCheckbox.enabled = cellsEnabled;
            roundCornerInput.enabled = cellsEnabled && roundCornerCheckbox.value;
            opacitySlider.enabled = cellsEnabled;
        }

        cellRectCheckbox.onClick = function () {
            updateCellOptionEnabled();
            safeUpdatePreview();
        };

        // 角丸トグルで数値欄の有効/無効を更新してプレビュー / Toggle round corners: refresh the field enable, then preview
        roundCornerCheckbox.onClick = function () {
            updateCellOptionEnabled();
            safeUpdatePreview();
        };
        changeValueByArrowKey(roundCornerInput);
        attachLivePreview(roundCornerInput);

        // OKボタン押下時（ドキュメントは変更せず閉じるだけ。クリア/描画は確定処理 finalAction で実行）
        // OK button: just close (no document mutation here — clear/draw happens in finalAction to keep undo bookkeeping correct)
        btnOK.onClick = function () {
            updateGutterEnabled();
            dialog.close(1);
        };

        // キャンセル：ダイアログを閉じるだけ（後処理は dialog.show() 後の分岐で rollback）/ Cancel: just close; cleanup/rollback happens after dialog.show()
        btnCancel.onClick = function () {
            dialog.close(0);
        };

        // 行・列・ガター・伸張・ガイド系の設定を収集 / Collect grid (rows/cols/gutters/extension/guide) settings
        function collectGridSettings() {
            return {
                columnCount: toInteger(columnCountInput.text, 0),
                rowCount: toInteger(rowCountInput.text, 0),
                guideExtension: (extensionCheckbox.value ? toNumber(extensionInput.text, 0) : 0) * unitFactor,
                rowGutter: toNumber(rowGutterInput.text, 0) * unitFactor,
                columnGutter: toNumber(columnGutterInput.text, 0) * unitFactor,
                drawGridGuides: drawGuidesCheckbox.value,
                drawArtboardEdge: artboardEdgeCheckbox.value,
                splitCells: splitCellCheckbox.value
            };
        }

        // 上下左右マージンを収集 / Collect top/bottom/left/right margins
        function collectMarginSettings() {
            return {
                marginTop: toNumber(marginTopInput.text, 0) * unitFactor,
                marginBottom: toNumber(marginBottomInput.text, 0) * unitFactor,
                marginLeft: toNumber(marginLeftInput.text, 0) * unitFactor,
                marginRight: toNumber(marginRightInput.text, 0) * unitFactor
            };
        }

        // セル（長方形化・角丸・不透明度）の設定を収集 / Collect cell (rect/round/opacity) settings
        function collectCellSettings() {
            return {
                drawCells: cellRectCheckbox.value,
                cornerRadius: roundCornerCheckbox.value ? toNumber(roundCornerInput.text, 0) * unitFactor : 0,
                cellOpacity: Math.round(opacitySlider.value)
            };
        }

        // 各 collector をまとめて描画コンテキストを構築 / Combine the collectors into a draw context
        function buildDrawContext(isPreview) {
            var drawContext = {
                doc: doc,
                isPreview: isPreview,
                allBoards: isAllArtboards(),
                selBounds: isSelectionMode() ? cachedSelectionBounds : null
            };
            mergeInto(drawContext, collectGridSettings());
            mergeInto(drawContext, collectMarginSettings());
            mergeInto(drawContext, collectCellSettings());
            return drawContext;
        }

        // 属性パネルの「中心を表示」を選択オブジェクトに適用 / Apply Attributes-panel "Show Center" to the current selection
        // API で直接設定できないため、一時アクション（.aia）を読み込んで再生 / No direct API, so load and play a temporary action
        function applyShowCenterAction() {
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

        /**
         * grid_guidesレイヤーのガイドを削除する。
         * 対象が「すべてのアートボード」なら全部、それ以外はアクティブなアートボード上のガイドのみ。
         * Remove guides from the grid_guides layer: all of them for "all artboards", otherwise only those on the active artboard.
         * @returns {boolean} 1つ以上削除したら true
         */
        function clearExistingGuides() {
            var guidesLayer = null;
            for (var i = 0; i < doc.layers.length; i++) {
                if (doc.layers[i].name === GUIDE_LAYER_NAME) {
                    guidesLayer = doc.layers[i];
                    break;
                }
            }
            if (!guidesLayer) return false;

            // すべてのアートボードが対象なら範囲を絞らない / No limit when every artboard is a target
            var limitRect = isAllArtboards()
                ? null
                : doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;

            safeUnlockLayer(guidesLayer);
            var removedCount = 0;
            for (var j = guidesLayer.pageItems.length - 1; j >= 0; j--) {
                var item = guidesLayer.pageItems[j];
                if (!item.guides) continue;
                if (limitRect && !isCenterInsideRect(item, limitRect)) continue;
                item.remove();
                removedCount++;
            }
            return removedCount > 0;
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
        updateGutterEnabled();
        updateCellOptionEnabled();
        syncLinkedMargins();
        updateTargetMode();
        safeUpdatePreview();

        if (dialog.show() === 1) {
            // OK: rollback preview and execute final drawing once so user can undo in one step
            previewManager.confirm(function () {
                // プレビューレイヤーを先に削除（最後に消すと選択が解除されるため）/ Remove preview layer first (removing it later clears the selection)
                safeRemoveLayerByName(doc, PREVIEW_LAYER_NAME);
                if (clearGuidesCheckbox.value) {
                    clearExistingGuides();
                }
                drawnCellItems = drawGrid(buildDrawContext(false));
                // 選択オブジェクトの処理 / Handle original selected objects
                if (cachedSelectionItems.length > 0 && isSelectionMode()) {
                    if (removeOriginalRadio && removeOriginalRadio.value) {
                        for (var i = cachedSelectionItems.length - 1; i >= 0; i--) {
                            var itemToRemove = cachedSelectionItems[i];
                            // ロック・非表示などで失敗しても他を続行 / Keep going even if one remove fails (locked/hidden, etc.)
                            safeExecute(function () { itemToRemove.remove(); });
                        }
                    } else if (makeOriginalGuidesRadio && makeOriginalGuidesRadio.value) {
                        // 元オブジェクトをガイド化 / Convert originals to guides
                        try {
                            doc.selection = null;
                            for (var j = 0; j < cachedSelectionItems.length; j++) {
                                cachedSelectionItems[j].selected = true;
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
                        for (var k = 0; k < drawnCellItems.length; k++) {
                            drawnCellItems[k].selected = true;
                        }
                        applyShowCenterAction();
                        doc.selection = null;
                    }
                } catch (e) {
                    $.writeln("[GenerateGuidesGrid] show center error: " + e);
                }
            });
        } else {
            // Cancel: rollback preview changes
            previewManager.rollback();
            // Cleanup preview layer just in case (fallback)
            safeRemoveLayerByName(doc, PREVIEW_LAYER_NAME);
            // プレビューで外れた選択を元に戻す / Restore selection lost during preview
            safeRestoreSelection();
        }
    }

    main();

})();
