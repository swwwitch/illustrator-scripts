#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ダイアログで方向・位置・単位・対象（カンバス／アートボード）を指定してガイドを作成します。

### 主な機能

- 水平／垂直ガイドの作成
- 位置と延長（アートボード外への伸ばし量）の指定
- 単位選択（in / mm / pt / px など）
- 環境設定に基づく初期単位の自動取得
- ↑↓キー（Shift で 10 単位スナップ）による数値調整
- リピート（本数・間隔）の指定
- ガイドのライブプレビュー
- ガイドは「_guide」レイヤーに作成してロック

### note

https://note.com/dtp_tranist/n/n1085336d7265

### 更新履歴

- v1.0 (20250713): 初期バージョン
- v1.2.1 (20260706): ダイアログウィンドウ生成を関数化、記号のロケール対応、各入力ラベルに「：」を付与、2カラム＋下部ボタンバー化、単位を共有ドロップダウン＋各欄の単位表記に集約、カンバスの原点補正を撤回（座標系不一致による誤配置を修正）、UIヘルパー（addPanel/addColumnGroup/addLabeledField）で重複整理、主要フィールドに tooltip 追加、未使用の記号ケースを整理

*/

/*

### Overview

Creates guides in Illustrator by specifying direction, position, unit and target (canvas or artboard) through a dialog.

### Features

- Horizontal / vertical guides
- Position and extension (beyond the artboard)
- Unit selection (in / mm / pt / px, etc.)
- Auto initial unit from preferences
- Up/Down keys (Shift snaps to 10) to adjust values
- Repeat (count / distance)
- Live guide preview
- Guides created on a locked "_guide" layer

### Flow

1. Show dialog
2. Enter settings
3. Create guides on OK

### Change Log

- v1.0 (20250713): Initial version
- v1.2.1 (20260706): Extracted dialog-window creation into a helper, added a locale-aware colon, appended colons to input labels, two-column layout with a bottom button bar, consolidated units into a shared dropdown + per-field unit labels, reverted canvas ruler-origin offset (fixed misplacement from coordinate-space mismatch), deduplicated UI with helpers (addPanel/addColumnGroup/addLabeledField), added tooltips to key fields, trimmed unused symbol cases

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "NewGuideMaker";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User configuration
// =========================================

/* カンバス端まで届く十分な長さ（Illustrator の最大カンバス 227inch 相当）/ Length long enough to span the canvas (227 inch ≈ Illustrator max canvas) */
var CANVAS_SPAN_PT = 227 * 72;

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 6;                  /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

// =========================================
// ローカライズ / Localization
// =========================================

/* ロケール取得関数 / Get current language */
function getCurrentLang() {
    var loc = ($.locale || "") + ""; // ensure string
    // Treat locales starting with "ja" (e.g., "ja", "ja_JP") as Japanese
    if (loc.indexOf("ja") === 0) {
        return "ja";
    }
    return "en";
}
var lang = getCurrentLang();

$.localize = true;

/* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
var LABELS = {
    dialog: {
        title: { ja: "ガイド作成", en: "Create Guide" }
    },
    target: {
        label:    { ja: "対象", en: "Target" },
        canvas:   { ja: "カンバス", en: "Canvas" },
        artboard: { ja: "アートボード", en: "Artboard" },
        extension: { ja: "延長", en: "Extension" }
    },
    direction: {
        label:      { ja: "方向", en: "Direction" },
        horizontal: { ja: "水平方向", en: "Horizontal" },
        vertical:   { ja: "垂直方向", en: "Vertical" },
        position:   { ja: "開始位置", en: "Start Position" }
    },
    layer: {
        label:   { ja: "作成レイヤー", en: "Target Layer" },
        guide:   { ja: "_guideレイヤー", en: "_guide Layer" },
        current: { ja: "現在のレイヤー", en: "Current Layer" }
    },
    repeat: {
        label:    { ja: "リピート", en: "Repeat" },
        count:    { ja: "ガイド数", en: "Guide Count" },
        distance: { ja: "距離", en: "Distance" }
    },
    unit: {
        label: { ja: "単位", en: "Unit" }
    },
    tip: {
        extension: { ja: "ガイドをアートボードの外側へ伸ばす量（アートボード対象時のみ）", en: "How far to extend guides beyond the artboard (artboard target only)" },
        position:  { ja: "ガイドの開始位置。↑↓で増減、Shift+↑↓で10単位スナップ", en: "Guide start position. Up/Down to step, Shift+Up/Down snaps to 10" },
        count:     { ja: "作成するガイドの本数", en: "Number of guides to create" },
        distance:  { ja: "リピート時のガイドの間隔", en: "Spacing between repeated guides" },
        direction: { ja: "H / V キーでも切り替えできます", en: "Toggle with the H / V keys too" }
    },
    button: {
        ok:     { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    alert: {
        locked: { ja: "アクティブレイヤーがロックされています。", en: "The active layer is locked." },
        noDoc:  { ja: "ドキュメントが開かれていません。", en: "No document is open." }
    }
};

/* ラベル取得（例: L('target','canvas')）/ Resolve a label (e.g. L('target','canvas')) */
function L() {
    var node = LABELS;
    for (var i = 0; i < arguments.length; i++) {
        if (node == null) break;
        node = node[arguments[i]];
    }
    return (node && node[lang] != null) ? node[lang] : "";
}

/* コロン記号（日本語は全角、英語は半角）/ Colon symbol (full-width for Japanese, half-width for English) */
function uiColon() {
    return (lang === "ja") ? "：" : ":";
}

// =========================================
// 単位 / Units
// =========================================

/* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; array index equals the rulerType code (0=in, 1=mm, 2=pt …) */
var UNITS = [
    { label: "in",    factor: 72.0 },                // 0
    { label: "mm",    factor: 72.0 / 25.4 },         // 1
    { label: "pt",    factor: 1.0 },                 // 2
    { label: "pica",  factor: 12.0 },                // 3
    { label: "cm",    factor: 72.0 / 2.54 },         // 4
    { label: "Q/H",   factor: 72.0 / 25.4 * 0.25 },  // 5
    { label: "px",    factor: 1.0 },                 // 6
    { label: "ft/in", factor: 72.0 * 12.0 },         // 7
    { label: "m",     factor: 72.0 / 25.4 * 1000.0 },// 8
    { label: "yd",    factor: 72.0 * 36.0 },         // 9
    { label: "ft",    factor: 72.0 * 12.0 }          // 10
];

/* 単位ラベルの一覧（ドロップダウン用）/ List of unit labels (for the dropdown) */
function unitLabels() {
    var labels = [];
    for (var i = 0; i < UNITS.length; i++) {
        labels.push(UNITS[i].label);
    }
    return labels;
}

/* ルーラー環境設定の単位インデックス（= rulerType コード）/ Current ruler unit index (= rulerType code) */
function rulerUnitIndex() {
    var code = app.preferences.getIntegerPreference("rulerType");
    return (code >= 0 && code < UNITS.length) ? code : 2; // 既定 pt / default pt
}

/* 値と単位ラベルからptに変換（数値以外は0扱い）/ Convert value + unit label to pt (non-numeric becomes 0) */
function convertToPt(value, label) {
    var numValue = Number(value);
    if (isNaN(numValue)) {
        return 0;
    }
    for (var i = 0; i < UNITS.length; i++) {
        if (UNITS[i].label === label) {
            return numValue * UNITS[i].factor;
        }
    }
    return numValue; // 見つからなければ pt 扱い / fall back to pt
}

// =========================================
// UIレイアウト補助 / UI layout helpers
// =========================================

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* 行グループの共通設定（ボタン列など） / Apply a horizontal row group */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    group.alignment = alignment || "left";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* ラベル付きパネルを生成（共通レイアウト適用）/ Add a titled panel with the shared layout */
function addPanel(parent, labelText) {
    var panel = parent.add("panel");
    panel.text = labelText;
    setupPanel(panel);
    return panel;
}

/* 左寄せの縦並びグループを生成（ラジオ列など）/ Add a left-aligned vertical group (e.g. radio column) */
function addColumnGroup(parent) {
    var group = parent.add("group");
    group.orientation = "column";
    group.alignChildren = ["left", "center"];
    return group;
}

/* ダイアログウィンドウを生成（上段に内容・下段にボタンの縦構成）/ Create a configured dialog window (content on top, buttons at the bottom) */
function createDialogWindow(title) {
    var win = new Window("dialog", title);
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];
    win.spacing = WINDOW_SPACING;
    win.margins = WINDOW_MARGINS;
    return win;
}

// =========================================
// ガイド座標 / Guide geometry
// =========================================

/* ガイド1本分の始点・終点座標を返す（カンバスはドキュメント原点基準＝既定のルーラー0点）/ Return start/end points for a single guide (canvas is measured from the document origin, i.e. the default ruler zero) */
function guidePathPoints(isCanvas, isHorizontal, positionPt, extensionPt, artboardRect) {
    if (isCanvas) {
        /* Y は上方向が正なので、下向きの位置は減算 / Y is up, so a downward position subtracts */
        return isHorizontal
            ? [[-CANVAS_SPAN_PT, -positionPt], [CANVAS_SPAN_PT, -positionPt]]
            : [[positionPt, CANVAS_SPAN_PT], [positionPt, -CANVAS_SPAN_PT]];
    }
    var left = artboardRect[0];
    var top = artboardRect[1];
    var right = artboardRect[2];
    var bottom = artboardRect[3];
    return isHorizontal
        ? [[left - extensionPt, top - positionPt], [right + extensionPt, top - positionPt]]
        : [[left + positionPt, top + extensionPt], [left + positionPt, bottom - extensionPt]];
}

// =========================================
// メイン / Main
// =========================================

(function () {
    if (app.documents.length === 0) {
        alert(L('alert', 'noDoc'));
        return;
    }
    var dialog = createGuideDialog();
    dialog.show();

    /* ダイアログの構築 / Build dialog UI */
    function createGuideDialog() {
        var doc = app.activeDocument;

        /* 単位オプションと現在単位のインデックス / Unit options and current unit index */
        var unitOptions = unitLabels();
        var currentUnitIndex = rulerUnitIndex();
        var unitTexts = []; /* 各数値欄の単位表記（共有ドロップダウンに追従）/ Per-field unit labels (follow the shared dropdown) */

        /* _guide レイヤーは使用時のみ遅延生成 / _guide layer is created lazily, only when used */
        var guideLayer = null;
        function ensureGuideLayer() {
            if (guideLayer) {
                return guideLayer;
            }
            var layers = doc.layers;
            for (var i = 0; i < layers.length; i++) {
                if (layers[i].name === "_guide") {
                    guideLayer = layers[i];
                    return guideLayer;
                }
            }
            guideLayer = layers.add();
            guideLayer.name = "_guide";
            return guideLayer;
        }

        /* 現在描画中のプレビュー線（常に配列で保持）/ Currently drawn preview paths (always an array) */
        var activePreviewPaths = null;

        /* プレビュー線を削除（モーダル中はドキュメント編集不可なのでパスは常に有効）/ Remove preview paths (paths stay valid during a modal dialog) */
        function removePreviewPaths() {
            if (!activePreviewPaths) return;
            for (var i = 0; i < activePreviewPaths.length; i++) {
                var path = activePreviewPaths[i];
                if (path && !path.locked && path.layer && !path.layer.locked) {
                    path.remove();
                }
            }
            activePreviewPaths = null;
        }

        /* プレビュー描画先レイヤーを決定（不可なら null）/ Resolve the layer to draw on (null if unavailable) */
        function resolvePreviewLayer() {
            if (guideLayerRadio.value) {
                var layer = ensureGuideLayer();
                layer.locked = false;
                return layer;
            }
            var active = doc.activeLayer;
            if (active.locked) {
                alert(L('alert', 'locked'));
                return null;
            }
            return active;
        }

        /* プレビュー線の見た目を設定 / Style a preview path */
        function stylePreviewPath(path) {
            var colorSpace = app.activeDocument.documentColorSpace;
            var color;
            if (colorSpace === DocumentColorSpace.CMYK) {
                color = new CMYKColor();
                color.cyan = 70;
                color.magenta = 50;
                color.yellow = 0;
                color.black = 0;
            } else {
                /* CMYK 以外は青の RGB にフォールバック / Fall back to a blue RGB for non-CMYK modes */
                color = new RGBColor();
                color.red = 74;
                color.green = 132;
                color.blue = 255;
            }
            path.stroked = true;
            path.filled = false;
            path.strokeWidth = 1.0;
            path.strokeColor = color;
            path.guides = false; /* プレビューはガイド化しない / Preview is not a guide */
        }

        /* プレビュー線を描画 / Draw preview paths */
        function drawPreview() {
            removePreviewPaths();

            /* 単位は全フィールド共通 / A single shared unit for all fields */
            var unit = unitDropdown.selection.text;

            var position = parseFloat(positionInput.text);
            if (isNaN(position)) {
                return;
            }
            var positionPt = convertToPt(position, unit);

            var repeatCount = parseInt(repeatCountInput.text, 10);
            if (isNaN(repeatCount) || repeatCount < 1) repeatCount = 1;
            var repeatDistancePt = convertToPt(repeatDistanceInput.text, unit);
            /* 距離0以下なら重複を避けて1本に / Avoid overlapping guides when distance is 0 or less */
            if (repeatDistancePt <= 0) {
                repeatDistancePt = 0;
                repeatCount = 1;
            }

            var targetLayer = resolvePreviewLayer();
            if (!targetLayer) {
                return;
            }

            /* アートボード対象時の延長量と矩形を一度だけ算出 / Compute extension amount and rect once when targeting the artboard */
            var isCanvas = canvasRadio.value;
            var extensionPt = 0;
            var artboardRect = null;
            if (!isCanvas) {
                var extension = parseFloat(extensionInput.text);
                if (isNaN(extension)) {
                    return;
                }
                extensionPt = convertToPt(extension, unit);
                artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            }

            var drawnPaths = [];
            for (var i = 0; i < repeatCount; i++) {
                var currentPositionPt = positionPt + i * repeatDistancePt;
                var path = targetLayer.pathItems.add();
                path.setEntirePath(guidePathPoints(isCanvas, horizontalRadio.value, currentPositionPt, extensionPt, artboardRect));
                stylePreviewPath(path);
                drawnPaths.push(path);
            }
            activePreviewPaths = drawnPaths;
            app.redraw();
        }

        /* 数値入力欄（↑↓ステップ＋逐次プレビュー）を生成。幅は既定2文字 / Add a numeric field (default 2-char width) with arrow stepping and live preview */
        function addNumberField(parent, defaultValue, chars) {
            var field = parent.add('edittext {characters: ' + ((typeof chars === "number") ? chars : 2) + '}');
            field.text = defaultValue;
            changeValueByArrowKey(field, drawPreview);
            field.addEventListener("changing", drawPreview);
            return field;
        }

        /* 数値欄の右に単位表記を追加（共有ドロップダウンに追従）/ Add a unit label next to a field (follows the shared dropdown) */
        function addUnitText(parent) {
            var label = parent.add("statictext", undefined, unitOptions[currentUnitIndex]);
            label.preferredSize.width = 34;
            unitTexts.push(label);
            return label;
        }

        /* ラベル付き数値欄を生成（単位表記・ツールチップ・幅は任意）/ Add a labeled numeric field (optional unit label, tooltip, width) */
        function addLabeledField(parent, labelText, defaultValue, withUnit, tip, chars) {
            var row = parent.add("group");
            setupRow(row, "left", 6);
            var label = row.add("statictext", undefined, labelText + uiColon());
            var input = addNumberField(row, defaultValue, chars);
            if (withUnit) {
                addUnitText(row);
            }
            if (tip) {
                label.helpTip = tip;
                input.helpTip = tip;
            }
            return { row: row, input: input };
        }

        /* 設定用カラムを生成 / Create a settings column */
        function addSettingsColumn(parent) {
            var column = parent.add("group");
            column.orientation = "column";
            column.alignChildren = ["fill", "top"];
            column.spacing = WINDOW_SPACING;
            return column;
        }

        var dialog = createDialogWindow(L('dialog', 'title') + ' ' + SCRIPT_VERSION);

        /* 上段：2カラムを横並びに収める行 / Top area: a row holding the two columns */
        var columnsRow = dialog.add("group");
        columnsRow.orientation = "row";
        columnsRow.alignChildren = ["fill", "top"];
        columnsRow.spacing = COLUMN_SPACING;

        /* 左カラム：対象・作成レイヤー・単位 / Left column: target, layer, unit */
        var leftColumn = addSettingsColumn(columnsRow);

        /* 対象選択パネル / Target selection panel */
        var targetPanel = addPanel(leftColumn, L('target', 'label'));
        var targetRadioGroup = addColumnGroup(targetPanel);
        var canvasRadio = targetRadioGroup.add("radiobutton", undefined, L('target', 'canvas'));
        var artboardRadio = targetRadioGroup.add("radiobutton", undefined, L('target', 'artboard'));
        artboardRadio.value = true;

        /* 延長：ガイドをアートボード外へ伸ばす量 / Extension: how far to extend guides beyond the artboard */
        var extensionField = addLabeledField(targetPanel, L('target', 'extension'), "0", true, L('tip', 'extension'));
        var extensionRow = extensionField.row;
        var extensionInput = extensionField.input;

        /* 作成レイヤーパネル / Target layer panel */
        var layerPanel = addPanel(leftColumn, L('layer', 'label'));
        var layerRadioGroup = addColumnGroup(layerPanel);
        var guideLayerRadio = layerRadioGroup.add("radiobutton", undefined, L('layer', 'guide'));
        var currentLayerRadio = layerRadioGroup.add("radiobutton", undefined, L('layer', 'current'));
        guideLayerRadio.value = true; /* デフォルト / default */

        /* 右カラム：方向・リピート / Right column: direction, repeat */
        var rightColumn = addSettingsColumn(columnsRow);

        /* 方向選択パネル / Direction selection panel */
        var directionPanel = addPanel(rightColumn, L('direction', 'label'));
        directionPanel.helpTip = L('tip', 'direction');
        var directionRadioGroup = addColumnGroup(directionPanel);
        var horizontalRadio = directionRadioGroup.add("radiobutton", undefined, L('direction', 'horizontal'));
        var verticalRadio = directionRadioGroup.add("radiobutton", undefined, L('direction', 'vertical'));
        horizontalRadio.value = true;

        /* 位置 / Position */
        var positionInput = addLabeledField(directionPanel, L('direction', 'position'), "0", true, L('tip', 'position')).input;

        /* リピートパネル / Repeat panel */
        var repeatPanel = addPanel(rightColumn, L('repeat', 'label'));
        var repeatColumn = addColumnGroup(repeatPanel);
        var repeatCountInput = addLabeledField(repeatColumn, L('repeat', 'count'), "1", false, L('tip', 'count')).input;
        var repeatDistanceInput = addLabeledField(repeatColumn, L('repeat', 'distance'), "0", true, L('tip', 'distance'), 3).input;

        /* ボタン領域（左右分割）：左＝単位／スペーサー／右＝キャンセル＋OK / Bottom bar (left/right split): unit (left) · spacer · Cancel + OK (right) */
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, 10, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        /* 左側グループ：単位 / Left-side group: unit */
        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.add("statictext", undefined, L('unit', 'label') + uiColon());
        var unitDropdown = btnLeftGroup.add("dropdownlist", undefined, unitOptions);
        unitDropdown.selection = currentUnitIndex;
        unitDropdown.onChange = function() {
            /* 各数値欄の単位表記を更新 / Update the per-field unit labels */
            var label = unitDropdown.selection.text;
            for (var i = 0; i < unitTexts.length; i++) {
                unitTexts[i].text = label;
            }
            drawPreview();
        };

        /* スペーサー（伸縮）/ Spacer (stretchable) */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        /* 右側グループ：キャンセル＋OK（Mac 規約で Cancel → OK）/ Right-side group: Cancel + OK (Cancel → OK per macOS) */
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        var btnCancel = btnRightGroup.add("button", undefined, L('button', 'cancel'), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, L('button', 'ok'), { name: "ok" });

        /* OKボタン：プレビュー線をガイド化 / OK: convert preview paths to guides */
        btnOK.onClick = function() {
            if (activePreviewPaths) {
                /* ガイド化のため対象レイヤーのロックを一時解除 / Temporarily unlock so paths can be converted */
                if (guideLayer && guideLayer.locked) {
                    guideLayer.locked = false;
                }
                for (var i = 0; i < activePreviewPaths.length; i++) {
                    var path = activePreviewPaths[i];
                    if (path && path.layer && !path.layer.locked) {
                        path.guides = true;
                        path.strokeWidth = 0.1;
                    }
                }
            }
            /* "_guide" レイヤーを使った場合はロック / Lock the "_guide" layer when it was used */
            if (guideLayer && guideLayer.name === "_guide") {
                guideLayer.locked = true;
            }
            activePreviewPaths = null;
            dialog.close();
        };

        /* キャンセル：プレビューを消して閉じる / Cancel: remove preview then close */
        btnCancel.onClick = function() {
            removePreviewPaths();
            dialog.close();
        };

        /* イベント：対象切替（カンバス選択時は延長行をディム）/ Events: target switch (dim the extension row when Canvas is selected) */
        canvasRadio.onClick = function() {
            extensionRow.enabled = false;
            drawPreview();
        };
        artboardRadio.onClick = function() {
            extensionRow.enabled = true;
            drawPreview();
        };

        /* イベント：その他はプレビュー再描画のみ / Events: others just redraw the preview */
        horizontalRadio.onClick = drawPreview;
        verticalRadio.onClick = drawPreview;
        guideLayerRadio.onClick = drawPreview;
        currentLayerRadio.onClick = drawPreview;

        /* H/Vキーで方向切り替え / Switch direction with H/V keys */
        dialog.addEventListener("keydown", function(event) {
            if (event.keyName == "H" || event.keyName == "h") {
                horizontalRadio.value = true;
                verticalRadio.value = false;
                drawPreview();
                event.preventDefault();
            } else if (event.keyName == "V" || event.keyName == "v") {
                horizontalRadio.value = false;
                verticalRadio.value = true;
                drawPreview();
                event.preventDefault();
            }
        });

        /* 初期状態：カンバス選択時は延長行をディム / Initial state: dim the extension row when Canvas is the target */
        extensionRow.enabled = artboardRadio.value;

        /* ダイアログ表示時に「位置」入力欄へフォーカス / Focus position input on dialog show */
        positionInput.active = true;
        /* 初回プレビュー / Initial preview */
        drawPreview();
        /* ダイアログ閉じたらプレビュー線を消す / Remove preview paths on dialog close */
        dialog.onClose = function() {
            removePreviewPaths();
        };
        return dialog;
    }

    /*
      edittext に上下キーで値を増減させるイベントを追加（コールバック対応版）/ Add up/down arrow-key value stepping to an edittext (with onUpdate callback)
    */
    function changeValueByArrowKey(inputField, onUpdate) {
        inputField.addEventListener("keydown", function(event) {
            var value = Number(inputField.text);
            if (isNaN(value)) return;
            if (event.keyName != "Up" && event.keyName != "Down") return;

            /* 修飾キーは event から読む（keyboardState は macOS で誤報あり）/ Read the modifier from event (keyboardState misreports on macOS) */
            var shiftPressed = event.shiftKey;
            if (shiftPressed === undefined) {
                shiftPressed = ScriptUI.environment.keyboardState.shiftKey;
            }

            if (shiftPressed) {
                /* Shift押下時は「10の倍数」スナップ / Snap to multiples of 10 when Shift is pressed */
                value = Math.round(value / 10) * 10 + (event.keyName == "Up" ? 10 : -10);
            } else {
                value += (event.keyName == "Up" ? 1 : -1);
            }

            event.preventDefault();
            inputField.text = value;
            if (typeof onUpdate === "function") {
                onUpdate(inputField.text);
            }
        });
    }
})();
