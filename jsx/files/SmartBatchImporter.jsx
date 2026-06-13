#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

SmartBatchImporter.jsx

### 概要

- 複数のIllustratorファイル（.ai/.svg/.eps）を一括で読み込み、取り込んだファイル（またはアートボード）ごとに1つのアートボードを作成して、全体が正方形に近くなるグリッドでキャンバス中央に整列配置するスクリプトです。
- 読み込み対象の選択、種別フィルター、アートボード単位読み込み、ガイドの取り込み、スケール、ファイル名ラベルなど多彩なオプションを備えています。

### 主な機能

- 開いているファイル、またはフォルダー指定による一括読み込み（種別：AI / SVG / EPS）
- 取り込んだファイル（またはアートボード）ごとに1つのアートボードを作成
- 全体が正方形に近くなるよう列数・行数を自動調整し、キャンバス中央へ整列（完了後に全体を表示）
- アートボード単位での読み込み（ロック・非表示も対象、元のアートボードサイズと相対位置を保持）
- 短いガイドの取り込み（カンバスの半分以上の長いガイドは除外）
- スケール（％）指定（アートボード枠・内容・線幅を同率で）
- ファイル名ラベルの追加（フォルダー指定時は既定OFF）
- カラースペース（RGB/CMYK）、サイズプリセット（カスタム／A4／フルHD／ラージカンバス、単位自動切替）
- 読み込み後にファイルを閉じる／保持（開いているファイル選択時のみ。未保存変更があれば確認）
- プログレスバーとキャンセル対応
- 日本語／英語インターフェース対応

### 紹介記事

https://note.com/dtp_tranist/n/n8180588e5630

### 更新履歴

- v1.0.0 (20250529) : 初期バージョン
- v1.2.0 (20260613) : 読み込み対象選択（開いているファイル／フォルダー指定＋種別フィルタ AI/SVG）、新しいドキュメント設定（カラースペース・サイズプリセット・単位自動切替・ラージカンバス）、アートボード単位読み込み（ロック・非表示対象／元サイズ保持）、スケール、ファイル名ラベル、読み込み後の動作
- v1.2.1 (20260613) : 種別フィルタに EPS を追加、ダイアログのパネル構成を2カラム化（読み込み対象／種別、カラースペース／ドキュメントサイズ、読み込みオプション）

---

### Script Name:

SmartBatchImporter.jsx

### Overview

- A script to batch import multiple Illustrator files (.ai/.svg/.eps), turn each imported file (or artboard) into one artboard, and arrange them into a near-square grid centered on the canvas.
- Offers source selection, a type filter, per-artboard import, guide import, scaling, filename labels, and more.

### Main Features

- Batch import from currently open files or a chosen folder (types: AI / SVG / EPS)
- One artboard per imported file (or artboard)
- Auto-adjusts the number of columns/rows so the whole layout is near-square, centered on the canvas (fits all in the window when done)
- Per-artboard import (includes locked/hidden objects, preserves the original artboard size and relative positions)
- Imports short guides (long guides half the canvas size or more are excluded)
- Scale (%) option (artboard frame, content, and stroke widths at the same rate)
- Adds file-name labels (off by default in folder mode)
- Color space (RGB/CMYK) and size presets (Custom / A4 / Full HD / Large Canvas, with auto unit switching)
- Close or keep source files after import (only when importing open files; confirms if there are unsaved changes)
- Progress bar with cancel support
- Japanese and English UI support

### Update History

- v1.0.0 (20250529): Initial version
- v1.2.0 (20260613): Source selection (open files / folder + AI/SVG type filter), new-document settings (color space, size presets, auto unit switching, Large Canvas), per-artboard import (incl. locked/hidden, preserves original size), scale, file-name label, and after-import action
- v1.2.1 (20260613): Added EPS to the type filter; reorganized the dialog into two-column panels (source / type, color space / document size, import options)
*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.2.1";

// =========================================
// ユーザー設定 / User Settings
// =========================================
var CONFIG = {
    spacingX: 100,                 // グループ間の横間隔（pt）/ Horizontal gap between groups (pt)
    spacingY: 100,                 // 行間の縦間隔（pt）/ Vertical gap between rows (pt)
    artboardMargin: 8.5,           // アートボードと内容の余白（pt）/ Margin around content within an artboard (pt)
    labelFont: "HiraginoSans-W3",  // ラベルのフォント / Font used for labels
    labelSize: 9,                  // ラベルの文字サイズ（pt）/ Label font size (pt)
    labelLayerName: "_label"       // ラベルを置くレイヤー名 / Name of the layer that holds labels
};

// =========================================
// ローカライズ / Localization
// =========================================

/* 現在のUI言語を返す（ja / en）/ Return the current UI language (ja / en) */
function getCurrentLanguage() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var currentLanguage = getCurrentLanguage();

var LABELS = {
    dialog: {
        title: { ja: "ファイル一括読み込み", en: "Batch Import Files" }
    },
    panel: {
        source: { ja: "読み込み対象", en: "Source" },
        newDocument: { ja: "新しいドキュメント", en: "New Document" },
        colorSpace: { ja: "カラースペース", en: "Color Space" },
        docSize: { ja: "ドキュメントサイズ", en: "Document Size" },
        options: { ja: "読み込みオプション", en: "Import Options" },
        afterImport: { ja: "読み込み後", en: "After Import" }
    },
    radio: {
        openFiles: { ja: "現在、開いているファイル", en: "Currently open files" },
        specifyFolder: { ja: "フォルダーを指定", en: "Specify folder" },
        rgb: { ja: "RGB", en: "RGB" },
        cmyk: { ja: "CMYK", en: "CMYK" },
        closeDoc: { ja: "閉じる", en: "Close" },
        keepOpen: { ja: "開いたまま", en: "Keep Open" }
    },
    checkbox: {
        byArtboard: { ja: "アートボード単位", en: "Import per artboard" },
        includeGuides: { ja: "短いガイドを含める", en: "Include short guides" },
        attachLabel: { ja: "ファイル名ラベルを追加", en: "Add file-name labels" },
        scale: { ja: "拡大・縮小", en: "Scale" }
    },
    preset: {
        custom: { ja: "カスタム", en: "Custom" },
        a4: { ja: "A4：210 × 297 mm", en: "A4: 210 × 297 mm" },
        fullHD: { ja: "フルHD：1920 × 1080 px", en: "Full HD: 1920 × 1080 px" },
        largeCanvas: { ja: "ラージカンバス", en: "Large Canvas" }
    },
    field: {
        width: { ja: "幅", en: "Width" },
        height: { ja: "高さ", en: "Height" },
        fileType: { ja: "種別", en: "Type" }
    },
    button: {
        specify: { ja: "指定", en: "Choose" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        ok: { ja: "OK", en: "OK" }
    },
    progress: {
        title: { ja: "ファイルを読み込み中...", en: "Importing files..." },
        count: { ja: "読み込み", en: "Imported" }
    },
    prompt: {
        selectFolder: {
            ja: "読み込むファイルが入ったフォルダーを選択してください",
            en: "Select a folder that contains files to import"
        }
    },
    alert: {
        noValidFile: {
            ja: "有効なファイルが見つかりませんでした。",
            en: "No valid files found."
        },
        invalidNumber: {
            ja: "数値が正しくありません。",
            en: "Invalid numeric input."
        },
        pasteFail: {
            ja: "ペーストに失敗しました",
            en: "Paste failed"
        },
        cancelled: {
            ja: "読み込みを中断しました。ここまでに読み込んだ内容は新規ドキュメントに残っています。",
            en: "Import was stopped. Items imported so far remain in the new document."
        }
    },
    confirm: {
        discardUnsaved: {
            ja: "開いているファイルに未保存の変更がある場合、「読み込み後に閉じる」を選ぶと保存せずに閉じます。続行しますか？",
            en: "If any open files have unsaved changes, choosing \"Close After Import\" closes them without saving. Continue?"
        }
    },
    tooltip: {
        openFiles: {
            ja: "現在Illustratorで開いているドキュメントを読み込み対象にします。",
            en: "Uses the documents currently open in Illustrator as the import source."
        },
        specifyFolder: {
            ja: "指定したフォルダー内の .ai / .svg ファイルを読み込み対象にします。",
            en: "Uses .ai / .svg files in the chosen folder as the import source."
        },
        fileType: {
            ja: "フォルダー指定時に読み込むファイル形式を選びます。",
            en: "Choose which file types to import in folder mode."
        },
        closeDoc: {
            ja: "読み込み後に元ファイルを保存せずに閉じます。未保存の変更があるファイルでは変更が失われます。",
            en: "Closes source files after import without saving. Unsaved changes in those files will be lost."
        },
        keepOpen: {
            ja: "読み込み後も元ファイルを開いたままにします。アートボード単位読み込みで一時的に解除したロック・非表示状態は復元します。",
            en: "Keeps source files open after import. Lock/hidden states temporarily changed for per-artboard import are restored."
        },
        specify: {
            ja: "読み込むファイル（.ai / .svg）が入ったフォルダーを選びます。",
            en: "Choose a folder that contains the files to import (.ai / .svg)."
        },
        colorSpace: {
            ja: "新規ドキュメントのカラースペースを選びます。読み込み元の色を完全に変換する機能ではありません。",
            en: "Choose the color space for the new document. This does not fully convert all colors from source files."
        },
        byArtboard: {
            ja: "各アートボードを別々に読み込み、元のアートボードサイズと相対位置を保ちます。ロック・非表示オブジェクトも対象です。",
            en: "Imports each artboard separately and preserves the original artboard size and relative positions. Locked/hidden objects are included."
        },
        includeGuides: {
            ja: "短いガイドも読み込みます。カンバスの半分以上の長いガイドは除外します。",
            en: "Also imports short guides. Long guides that are half the canvas size or more are excluded."
        },
        attachLabel: {
            ja: "読み込んだアートボードの下に、元ファイル名のラベルを追加します。",
            en: "Adds a source file-name label below each imported artboard."
        },
        scale: {
            ja: "読み込む内容とアートボード枠を指定％で拡大・縮小します。線幅も同率で変わります。",
            en: "Scales imported content and artboard cells by the specified percentage. Stroke widths scale by the same ratio."
        },
        size: {
            ja: "新規ドキュメントの初期サイズです。幅・高さの単位は選択したプリセットに連動します（カスタムは px）。",
            en: "Initial size of the new document. Width/height units follow the selected preset (px for Custom)."
        }
    }
};

/* ネストしたキー（"category.key"）からローカライズ文字列を取得する。
   現在言語が無ければ 英語 → 日本語 → path の順にフォールバックする。
   / Resolve a localized string from a dotted "category.key" path.
   Falls back current language → English → Japanese → the path itself. */
function getLocalizedText(path) {
    var parts = path.split('.');
    var category = LABELS[parts[0]];
    var entry = category ? category[parts[1]] : null;
    if (!entry) return path;
    var text = entry[currentLanguage];
    if (text === undefined) text = entry.en;
    if (text === undefined) text = entry.ja;
    if (text === undefined) return path;
    return text.replace(/\{slash\}/g, '/');
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(path) {
    return getLocalizedText(path) + (currentLanguage === 'ja' ? '：' : ':');
}

// =========================================
// 単位 / Units
// =========================================

// 1mm あたりのポイント数 / Points per millimeter
var MM_TO_PT = 2.8346;

// 各単位 → ポイントの換算係数 / Conversion factor from each unit to points
var UNIT_TO_PT = {
    px: 0.75,     // 96dpi: 1px = 0.75pt
    mm: MM_TO_PT, // 1mm = 2.8346pt
    inch: 72      // 1inch = 72pt
};

// ドキュメントサイズプリセット（値はそれぞれのネイティブ単位）/ Document size presets (values in their native unit)
// インデックスは presetDropdown の並びと一致。カスタムは unit のみ（値は自動入力しない）。
var SIZE_PRESETS = [
    { unit: "px" },                             // カスタム / Custom
    { unit: "mm", width: 210, height: 297 },    // A4
    { unit: "px", width: 1920, height: 1080 },  // フルHD / Full HD
    { unit: "inch", width: 2270, height: 2270 } // ラージカンバス / Large Canvas
];

/* パネルの余白と間隔 / Panel margins and spacing */
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 8;

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

/* ラベルを指定位置（アートボード下端の左下）に配置する / Place the label at the given bottom-left point (just below the artboard) */
function positionLabelBelow(label, left, bottomY) {
    label.left = left;
    label.top = bottomY - 4; // アートボード下端のすぐ下 / Just below the artboard's bottom edge
}

/* "_label" レイヤーを取得し、無ければ作成する / Get the "_label" layer, creating it if missing */
function getOrCreateLabelLayer(doc) {
    var layer;
    try {
        layer = doc.layers.getByName(CONFIG.labelLayerName);
    } catch (e) {
        layer = doc.layers.add();
        layer.name = CONFIG.labelLayerName;
    }
    return layer;
}

/* レイヤーとそのサブレイヤーのロック・非表示を再帰的に解除する / Recursively unlock and reveal a layer and its sublayers */
function unlockLayerTree(layer) {
    layer.locked = false;
    layer.visible = true;
    for (var i = 0; i < layer.layers.length; i++) {
        unlockLayerTree(layer.layers[i]);
    }
}

/* ドキュメント内のすべてのレイヤー・オブジェクトのロックと非表示を解除する / Unlock and reveal every layer and object in the document */
function unlockAllLayersAndItems(doc) {
    for (var li = 0; li < doc.layers.length; li++) {
        unlockLayerTree(doc.layers[li]);
    }
    for (var pi = 0; pi < doc.pageItems.length; pi++) {
        doc.pageItems[pi].locked = false;
        doc.pageItems[pi].hidden = false;
    }
}

/* レイヤー・オブジェクトのロック/非表示状態を記録する（後で元に戻すため）
   / Capture the lock/hidden state of layers and items so it can be restored later */
function captureLockHiddenState(doc) {
    var layers = [];
    var items = [];
    function walkLayers(layerList) {
        for (var i = 0; i < layerList.length; i++) {
            var lyr = layerList[i];
            layers.push({ ref: lyr, locked: lyr.locked, visible: lyr.visible });
            walkLayers(lyr.layers);
        }
    }
    walkLayers(doc.layers);
    for (var pi = 0; pi < doc.pageItems.length; pi++) {
        var it = doc.pageItems[pi];
        items.push({ ref: it, locked: it.locked, hidden: it.hidden });
    }
    return { layers: layers, items: items };
}

/* 記録したロック/非表示状態を元に戻す。1件失敗しても残りは復元を続ける。
   / Restore a previously captured lock/hidden state. A failure on one item won't stop the rest. */
function restoreLockHiddenState(state) {
    for (var i = 0; i < state.layers.length; i++) {
        try { state.layers[i].ref.locked = state.layers[i].locked; } catch (eLayerLock) { }
        try { state.layers[i].ref.visible = state.layers[i].visible; } catch (eLayerVis) { }
    }
    for (var j = 0; j < state.items.length; j++) {
        try { state.items[j].ref.locked = state.items[j].locked; } catch (eItemLock) { }
        try { state.items[j].ref.hidden = state.items[j].hidden; } catch (eItemHidden) { }
    }
}

/* 2つの矩形（[L, T, R, B]、T>B）が重なるか / Whether two [L, T, R, B] rects (T > B) overlap */
function rectsIntersect(a, b) {
    return !(a[2] < b[0] || a[0] > b[2] || a[3] > b[1] || a[1] < b[3]);
}

/* 取り込み対象のガイドを集める。長さがカンバスの半分以上のものは無視する。
   ロック・非表示のガイドは選択できないため除外。withinRect を渡すとその矩形に重なるものだけにする。
   / Collect importable guides, ignoring ones as long as half the canvas or more.
   Locked/hidden guides can't be selected, so they're excluded. With withinRect, keep only overlapping ones. */
function collectImportableGuides(doc, canvasWidth, canvasHeight, withinRect) {
    var result = [];
    var halfWidth = canvasWidth / 2;
    var halfHeight = canvasHeight / 2;
    for (var i = 0; i < doc.pageItems.length; i++) {
        var item = doc.pageItems[i];
        if (!item.guides || item.locked || item.hidden) continue;
        var gb = item.geometricBounds; // [L, T, R, B]
        if ((gb[2] - gb[0]) >= halfWidth || (gb[1] - gb[3]) >= halfHeight) continue; // 長すぎるガイドは無視
        if (withinRect && !rectsIntersect(gb, withinRect)) continue;
        result.push(item);
    }
    return result;
}

/* 複数オブジェクトの合成 visibleBounds を返す / Return the combined visibleBounds of multiple items */
function getCombinedVisibleBounds(items) {
    var b = items[0].visibleBounds;
    var left = b[0], top = b[1], right = b[2], bottom = b[3];
    for (var i = 1; i < items.length; i++) {
        var v = items[i].visibleBounds;
        if (v[0] < left) left = v[0];
        if (v[1] > top) top = v[1];
        if (v[2] > right) right = v[2];
        if (v[3] < bottom) bottom = v[3];
    }
    return [left, top, right, bottom];
}

/* 貼り付け直後のオブジェクトを1グループにまとめ、セル（＝追加するアートボード枠）単位で行送り配置し、
   セルに合わせたアートボードを追加、必要ならラベルを付ける
   / Group the just-pasted objects, lay them out cell-by-cell (one artboard per cell),
   add a matching artboard, and optionally a label.
   artboardCell を渡すと元のアートボードサイズと内容の相対位置を保持する。
   / When artboardCell is provided, the original artboard size and the content's relative position are preserved. */
function placePastedGroup(ctx, labelName, artboardCell) {
    var newDoc = ctx.newDoc;
    var pastedItems = newDoc.selection;
    if (pastedItems.length === 0) {
        alert(getLocalizedText('alert.pasteFail') + "：" + labelName);
        return false;
    }

    var targetLayer = newDoc.activeLayer;
    var pastedGroup = targetLayer.groupItems.add();
    for (var m = pastedItems.length - 1; m >= 0; m--) {
        pastedItems[m].moveToBeginning(pastedGroup);
    }
    newDoc.activeLayer = targetLayer; // 貼り付け直後にアクティブレイヤーを戻す / Restore the active layer right after pasting

    // スケール適用（コンテンツとアートボード枠の両方を同率で）。線幅・パターン・グラデーションも同率で縮める。
    // Apply scaling to both the content and the artboard cell at the same rate; also scale stroke widths / patterns / gradients.
    if (ctx.scalePercent !== 100) {
        pastedGroup.resize(
            ctx.scalePercent, ctx.scalePercent,
            true,             // changePositions
            true,             // changeFillPatterns
            true,             // changeFillGradients
            true,             // changeStrokePattern
            ctx.scalePercent  // changeLineWidths（線幅も同率で / scale stroke widths by the same percent）
        );
    }
    var scaleFactor = ctx.scalePercent / 100;

    var bounds = pastedGroup.visibleBounds;
    var contentWidth = bounds[2] - bounds[0];
    var contentHeight = bounds[1] - bounds[3];

    // セル（＝追加するアートボードの枠）の寸法と、その中での内容の位置
    // Cell (= the artboard to add) size and the content's position within it
    var pad = ctx.artboardPadding;
    var cellWidth, cellHeight, offsetX, offsetY;
    if (artboardCell) {
        // 元のアートボードサイズと相対位置を保持（スケールも反映）/ Preserve the original artboard size and relative position (scaled)
        cellWidth = artboardCell.width * scaleFactor;
        cellHeight = artboardCell.height * scaleFactor;
        offsetX = artboardCell.offsetX * scaleFactor;
        offsetY = artboardCell.offsetY * scaleFactor;
    } else {
        // 内容に余白を足した枠 / A box fitted to the content plus padding
        cellWidth = contentWidth + pad * 2;
        cellHeight = contentHeight + pad * 2;
        offsetX = pad;
        offsetY = pad;
    }

    // 内容の現在位置を基準に、相対オフセットを保ったままアートボードを作成（最終位置は後でグリッド配置）
    // Create the artboard around the content at its current spot, preserving the relative offset (final position is set later by the grid layout)
    var abLeft = bounds[0] - offsetX;
    var abTop = bounds[1] + offsetY;
    var newArtboard = newDoc.artboards.add([abLeft, abTop, abLeft + cellWidth, abTop - cellHeight]);
    newArtboard.name = labelName;

    var labelItem = null;
    if (ctx.showLabel) {
        var labelLayer = getOrCreateLabelLayer(newDoc);
        labelItem = newDoc.textFrames.add();
        labelItem.contents = labelName;
        // フォントが見つからない環境では既定フォントのまま / Keep the default font when the configured one is missing
        try {
            labelItem.textRange.characterAttributes.textFont = app.textFonts.getByName(CONFIG.labelFont);
        } catch (fontError) { }
        labelItem.textRange.characterAttributes.size = CONFIG.labelSize;
        // ラベルはアートボードの下端の左下に置く / Place the label below the artboard's bottom-left
        positionLabelBelow(labelItem, abLeft, abTop - cellHeight);
        if (labelItem.layer != labelLayer) labelItem.layer = labelLayer;
        labelItem.move(labelLayer, ElementPlacement.PLACEATBEGINNING);
    }

    // セルを記録（グループ・アートボード・ラベルを最終グリッド配置で一緒に動かす）
    // Record the cell so the group, artboard, and label can be moved together by the final grid layout
    ctx.cells.push({
        group: pastedGroup,
        artboard: newArtboard,
        label: labelItem,
        width: cellWidth,
        height: cellHeight
    });
    ctx.placedCount++;
    return true;
}

/* 記録したセルを「全体が正方形に近い」グリッドに並べ、キャンバス中央に配置する。
   セルごとに group・artboard・label を同じ移動量でまとめて動かす。
   / Lay out the recorded cells in a grid whose overall shape is as square as possible, centered on the canvas.
   For each cell, the group, artboard, and label move together by the same delta. */
function layoutCellsAsCenteredGrid(ctx) {
    var cells = ctx.cells;
    var count = cells.length;
    if (count === 0) return;

    // スロットサイズ＝最大セル寸法（サイズ混在でも重ならないように）/ Slot size = the largest cell (so mixed sizes never overlap)
    var slotWidth = 0;
    var slotHeight = 0;
    for (var i = 0; i < count; i++) {
        if (cells[i].width > slotWidth) slotWidth = cells[i].width;
        if (cells[i].height > slotHeight) slotHeight = cells[i].height;
    }

    // セル間の間隔。アートボードモードは「スロット幅の1/8」（縦も同値）、それ以外は CONFIG の既定値。
    // Gap between cells. In artboard mode use 1/8 of the slot width (same for vertical); otherwise the CONFIG default.
    var gapX = ctx.byArtboard ? (slotWidth / 8) : CONFIG.spacingX;
    var gapY = ctx.byArtboard ? (slotWidth / 8) : CONFIG.spacingY;

    // 全体のバウンディングボックスが最も正方形に近くなる列数を選ぶ
    // Choose the column count that makes the overall bounding box closest to square
    var columns = 1;
    var bestDiff = -1;
    for (var c = 1; c <= count; c++) {
        var r = Math.ceil(count / c);
        var totalWidthCandidate = c * slotWidth + (c - 1) * gapX;
        var totalHeightCandidate = r * slotHeight + (r - 1) * gapY;
        var diff = Math.abs(totalWidthCandidate - totalHeightCandidate);
        if (bestDiff < 0 || diff < bestDiff) {
            bestDiff = diff;
            columns = c;
        }
    }
    var rows = Math.ceil(count / columns);

    var totalWidth = columns * slotWidth + (columns - 1) * gapX;
    var totalHeight = rows * slotHeight + (rows - 1) * gapY;

    // グリッド全体をキャンバス中央に揃える / Center the whole grid on the canvas
    var startLeft = ctx.canvasCenterX - totalWidth / 2;
    var startTop = ctx.canvasCenterY + totalHeight / 2;

    for (var k = 0; k < count; k++) {
        var col = k % columns;
        var row = Math.floor(k / columns);
        var slotLeft = startLeft + col * (slotWidth + gapX);
        var slotTop = startTop - row * (slotHeight + gapY);

        var cell = cells[k];
        // セルをスロット内で中央に / Center the cell within its slot
        var targetLeft = slotLeft + (slotWidth - cell.width) / 2;
        var targetTop = slotTop - (slotHeight - cell.height) / 2;

        var rect = cell.artboard.artboardRect; // [L, T, R, B]
        var dx = targetLeft - rect[0];
        var dy = targetTop - rect[1];

        cell.group.translate(dx, dy);
        cell.artboard.artboardRect = [rect[0] + dx, rect[1] + dy, rect[2] + dx, rect[3] + dy];
        if (cell.label) cell.label.translate(dx, dy);
    }
}

/* メイン処理：ダイアログを表示し、選択されたソースを新規ドキュメントへ取り込み配置する
   / Main entry point: show the dialog, then import and arrange the chosen source into a new document */
(function main() {
    // 現在開いているドキュメントを収集 / Collect currently open documents
    var openDocs = [];
    for (var i = 0; i < app.documents.length; i++) {
        openDocs.push(app.documents[i]);
    }
    // ［指定］ボタンで選んだフォルダと、その中の対象ファイル一覧
    // The folder chosen via the [Choose] button and the matching files inside it
    var selectedFolder = null;
    var folderFiles = [];

    var dialog = new Window("dialog", getLocalizedText('dialog.title') + ' ' + SCRIPT_VERSION);
    dialog.orientation = "column";
    dialog.alignChildren = ["left", "top"];

    // --- 読み込み対象パネルと種別パネルを2カラムで並べる / Source panel and type panel in two columns ---
    var sourceRow = dialog.add("group");
    setupGroup(sourceRow, "row");
    sourceRow.alignChildren = ["left", "fill"]; // 読み込み対象と種別の高さを揃える / Match the two panels' heights

    // 左カラム：読み込み対象（ソース選択、ラジオは縦並び）/ Left column: source panel (radios stacked)
    var sourcePanel = sourceRow.add("panel", undefined, getLocalizedText('panel.source'));
    setupPanel(sourcePanel);
    var openDocsRadio = sourcePanel.add("radiobutton", undefined, getLocalizedText('radio.openFiles'));
    openDocsRadio.helpTip = getLocalizedText('tooltip.openFiles');

    var folderRow = sourcePanel.add("group");
    setupGroup(folderRow, "row");
    var folderRadio = folderRow.add("radiobutton", undefined, getLocalizedText('radio.specifyFolder'));
    folderRadio.helpTip = getLocalizedText('tooltip.specifyFolder');
    var selectFolderBtn = folderRow.add("button", undefined, getLocalizedText('button.specify'));
    selectFolderBtn.helpTip = getLocalizedText('tooltip.specify');

    // 選択したフォルダ名は別行に表示 / Show the chosen folder name on its own row
    var folderNameRow = sourcePanel.add("group");
    setupGroup(folderNameRow, "row");
    var folderNameText = folderNameRow.add("statictext", undefined, "", { truncate: "middle" });
    folderNameText.preferredSize = [260, 20];
    folderNameText.minimumSize = [260, 20];

    // 右カラム：種別（フォルダー指定時のフィルタ、縦並び）/ Right column: file types (folder-mode filter, stacked)
    var typePanel = sourceRow.add("panel", undefined, getLocalizedText('field.fileType'));
    setupPanel(typePanel);
    var aiCheckbox = typePanel.add("checkbox", undefined, "AI");
    aiCheckbox.helpTip = getLocalizedText('tooltip.fileType');
    var svgCheckbox = typePanel.add("checkbox", undefined, "SVG");
    svgCheckbox.helpTip = getLocalizedText('tooltip.fileType');
    var epsCheckbox = typePanel.add("checkbox", undefined, "EPS");
    epsCheckbox.helpTip = getLocalizedText('tooltip.fileType');
    aiCheckbox.value = true;   // 既定は AI のみ ON / Default: AI only
    svgCheckbox.value = false;
    epsCheckbox.value = false;

    // ファイル数はパネルのタイトルに「読み込み対象（5）」のように表示する
    // Show the file count in the panel title, e.g. "Source (5)"
    function updateFileCount() {
        var count = folderRadio.value ? folderFiles.length : openDocs.length;
        sourcePanel.text = getLocalizedText('panel.source') + (currentLanguage === 'ja' ? '（' + count + '）' : ' (' + count + ')');
    }

    // 「読み込み後の動作」は開いているファイルを選んだときのみ有効（フォルダ指定ではディム）
    // The "After Import" action is available only for open files (dimmed for folder import)
    function updateAfterImportState() {
        var enabled = openDocsRadio.value;
        afterImportLabel.enabled = enabled;
        closeRadio.enabled = enabled;
        keepOpenRadio.enabled = enabled;
    }

    // 種別は「フォルダー指定」のときのみ有効 / File types are available only in folder mode
    function updateTypeRowState() {
        typePanel.enabled = folderRadio.value;
    }

    // ファイル名ラベルはフォルダー指定では既定OFF、開いているファイルではON
    // Default the file-name label off in folder mode and on for open files
    function updateLabelOptionForMode() {
        showLabelCheckbox.value = openDocsRadio.value;
    }

    // 選択された種別の拡張子リスト / Extensions for the currently selected file types
    function getSelectedExtensions() {
        var exts = [];
        if (aiCheckbox.value) exts.push("ai");
        if (svgCheckbox.value) exts.push("svg");
        if (epsCheckbox.value) exts.push("eps");
        return exts;
    }

    // 選択中フォルダを現在の種別で再フィルタ / Re-filter the chosen folder by the current file types
    function refreshFolderFiles() {
        folderFiles = [];
        var exts = getSelectedExtensions();
        if (selectedFolder && exts.length > 0) {
            var typeRe = new RegExp("\\.(" + exts.join("|") + ")$", "i");
            folderFiles = selectedFolder.getFiles(function (f) {
                return f instanceof File && typeRe.test(f.name);
            }).sort(function (a, b) {
                return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1;
            });
        }
        updateFileCount();
    }

    aiCheckbox.onClick = function () {
        refreshFolderFiles();
    };
    svgCheckbox.onClick = aiCheckbox.onClick;
    epsCheckbox.onClick = aiCheckbox.onClick;

    // 開いているファイルが無ければフォルダ指定をデフォルトに
    // Default to folder mode when no document is open
    openDocsRadio.enabled = openDocs.length > 0;
    if (openDocs.length > 0) {
        openDocsRadio.value = true;
    } else {
        folderRadio.value = true;
    }

    // 2つのラジオは別コンテナにあるため、排他選択を手動で制御する
    // The two radios live in different containers, so enforce mutual exclusivity manually
    openDocsRadio.onClick = function () {
        folderRadio.value = false;
        updateFileCount();
        updateAfterImportState();
        updateTypeRowState();
        updateLabelOptionForMode();
    };
    folderRadio.onClick = function () {
        openDocsRadio.value = false;
        updateFileCount();
        updateAfterImportState();
        updateTypeRowState();
        updateLabelOptionForMode();
    };

    selectFolderBtn.onClick = function () {
        var folder = Folder.selectDialog(getLocalizedText('prompt.selectFolder'));
        if (!folder) return;
        selectedFolder = folder;
        folderRadio.value = true;
        openDocsRadio.value = false;
        var folderPath = decodeURI(folder.fsName);
        folderNameText.text = folderPath;
        folderNameText.helpTip = folderPath;
        refreshFolderFiles();
        updateAfterImportState();
        updateTypeRowState();
        updateLabelOptionForMode();
    };

    updateFileCount();
    updateTypeRowState();

    // --- 新しいドキュメントパネル（カラースペース＋ドキュメントサイズ）/ New document panel (color space + document size) ---
    var newDocPanel = dialog.add("panel", undefined, getLocalizedText('panel.newDocument'));
    setupPanel(newDocPanel);

    // カラースペースとドキュメントサイズを横2カラムで並べる / Lay out color space and document size in two columns
    var newDocColumns = newDocPanel.add("group");
    setupGroup(newDocColumns, "row");
    newDocColumns.alignChildren = ["left", "fill"]; // 2つのパネルの高さを揃える / Match the two panels' heights
    newDocColumns.spacing = 15;                     // カラースペースとサイズの2カラム間隔 / Gap between the color-space and size columns

    // カラースペース（ラジオは縦並び）/ Color space (radios stacked vertically)
    var colorPanel = newDocColumns.add("panel", undefined, getLocalizedText('panel.colorSpace'));
    setupPanel(colorPanel);
    var rgbRadio = colorPanel.add("radiobutton", undefined, getLocalizedText('radio.rgb'));
    rgbRadio.helpTip = getLocalizedText('tooltip.colorSpace');
    var cmykRadio = colorPanel.add("radiobutton", undefined, getLocalizedText('radio.cmyk'));
    cmykRadio.helpTip = getLocalizedText('tooltip.colorSpace');
    rgbRadio.value = true;

    var sizePanel = newDocColumns.add("panel", undefined, getLocalizedText('panel.docSize'));
    setupPanel(sizePanel);

    var presetDropdown = sizePanel.add("dropdownlist", undefined, [
        getLocalizedText('preset.custom'),
        getLocalizedText('preset.a4'),
        getLocalizedText('preset.fullHD'),
        getLocalizedText('preset.largeCanvas')
    ]);
    presetDropdown.selection = 0;

    // 選択中の単位（プリセット連動、カスタムは px）/ Current unit (follows the preset; px for Custom)
    var currentUnit = "px";

    // 幅と高さはそれぞれ別の行に / Width and height each on their own row
    var widthRow = sizePanel.add("group");
    setupGroup(widthRow, "row");
    var widthLabel = widthRow.add("statictext", undefined, labelText('field.width'));
    widthLabel.preferredSize = [40, 20];
    var widthInput = widthRow.add("edittext", undefined, "1000");
    widthInput.characters = 4;
    widthInput.helpTip = getLocalizedText('tooltip.size');
    var widthUnitText = widthRow.add("statictext", undefined, currentUnit);
    widthUnitText.preferredSize = [36, 20];

    var heightRow = sizePanel.add("group");
    setupGroup(heightRow, "row");
    var heightLabel = heightRow.add("statictext", undefined, labelText('field.height'));
    heightLabel.preferredSize = [40, 20];
    var heightInput = heightRow.add("edittext", undefined, "1000");
    heightInput.characters = 4;
    heightInput.helpTip = getLocalizedText('tooltip.size');
    var heightUnitText = heightRow.add("statictext", undefined, currentUnit);
    heightUnitText.preferredSize = [36, 20];

    // 入力値を旧単位から新単位へ換算する（物理サイズを保つ）/ Convert an input value from the old unit to the new one (keeps physical size)
    function convertUnitText(textValue, fromUnit, toUnit) {
        var value = parseFloat(textValue);
        if (isNaN(value)) return textValue;
        return String(Math.round(value * UNIT_TO_PT[fromUnit] / UNIT_TO_PT[toUnit]));
    }

    presetDropdown.onChange = function () {
        var preset = SIZE_PRESETS[presetDropdown.selection.index];
        var newUnit = preset.unit;
        if (preset.width !== undefined) {
            // プリセット：ネイティブ単位の値をそのまま入力 / Preset: set its native-unit values
            widthInput.text = String(preset.width);
            heightInput.text = String(preset.height);
        } else {
            // カスタム：現在の値を新しい単位へ換算して物理サイズを保つ / Custom: convert current values to keep the physical size
            widthInput.text = convertUnitText(widthInput.text, currentUnit, newUnit);
            heightInput.text = convertUnitText(heightInput.text, currentUnit, newUnit);
        }
        currentUnit = newUnit;
        widthUnitText.text = currentUnit;
        heightUnitText.text = currentUnit;
    };

    // --- 読み込みオプションパネル / Import options panel ---
    var optionsPanel = dialog.add("panel", undefined, getLocalizedText('panel.options'));
    setupPanel(optionsPanel);
    // オプションを2カラムで並べる（左：アートボード単位／ファイル名ラベル、右：短いガイド／拡大・縮小）
    // Lay out options in two columns (left: per-artboard / file-name label, right: short guides / scale)
    var optionsColumns = optionsPanel.add("group");
    setupGroup(optionsColumns, "row");
    optionsColumns.alignChildren = ["left", "top"]; // 2カラムを上端で揃える / Top-align the two columns
    optionsColumns.spacing = 20;                     // 左右カラムの間隔 / Gap between the two columns

    // 左カラム：アートボード単位／ファイル名ラベル / Left column
    var optionsLeft = optionsColumns.add("group");
    setupGroup(optionsLeft, "column");
    var byArtboardCheckbox = optionsLeft.add("checkbox", undefined, getLocalizedText('checkbox.byArtboard'));
    byArtboardCheckbox.value = true;
    byArtboardCheckbox.helpTip = getLocalizedText('tooltip.byArtboard');

    var showLabelCheckbox = optionsLeft.add("checkbox", undefined, getLocalizedText('checkbox.attachLabel'));
    showLabelCheckbox.value = openDocsRadio.value; // フォルダー指定では既定OFF / Default off in folder mode
    showLabelCheckbox.helpTip = getLocalizedText('tooltip.attachLabel');

    // 右カラム：短いガイド／拡大・縮小 / Right column
    var optionsRight = optionsColumns.add("group");
    setupGroup(optionsRight, "column");
    var includeGuidesCheckbox = optionsRight.add("checkbox", undefined, getLocalizedText('checkbox.includeGuides'));
    includeGuidesCheckbox.value = false;
    includeGuidesCheckbox.helpTip = getLocalizedText('tooltip.includeGuides');

    // スケール（チェックON時に％で拡大縮小）/ Scale (resize by percent when checked)
    var scaleRow = optionsRight.add("group");
    setupGroup(scaleRow, "row");
    var scaleCheckbox = scaleRow.add("checkbox", undefined, getLocalizedText('checkbox.scale'));
    scaleCheckbox.value = false;
    scaleCheckbox.helpTip = getLocalizedText('tooltip.scale');
    var scaleInput = scaleRow.add("edittext", undefined, "100");
    scaleInput.helpTip = getLocalizedText('tooltip.scale');
    scaleInput.characters = 5;
    scaleInput.enabled = scaleCheckbox.value;
    scaleRow.add("statictext", undefined, "%");
    scaleCheckbox.onClick = function () {
        scaleInput.enabled = this.value;
    };

    // 読み込み後の動作（ラベル＋ラジオの横並び。開いているファイル選択時のみ有効）
    // After-import action (label + radios in a row; enabled only when importing open files)
    var afterImportRow = optionsPanel.add("group");
    setupGroup(afterImportRow, "row");
    var afterImportLabel = afterImportRow.add("statictext", undefined, labelText('panel.afterImport'));
    var closeRadio = afterImportRow.add("radiobutton", undefined, getLocalizedText('radio.closeDoc'));
    closeRadio.helpTip = getLocalizedText('tooltip.closeDoc');
    var keepOpenRadio = afterImportRow.add("radiobutton", undefined, getLocalizedText('radio.keepOpen'));
    keepOpenRadio.helpTip = getLocalizedText('tooltip.keepOpen');
    closeRadio.value = true;

    updateAfterImportState(); // 初期状態を反映 / Apply the initial enabled/dimmed state

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";
    // name を "cancel"/"ok" にすると、クリックでダイアログが閉じる（Esc/Enter にも対応）
    // Naming them "cancel"/"ok" makes clicks dismiss the dialog (and binds Esc/Enter)
    var cancelBtn = buttonGroup.add("button", undefined, getLocalizedText('button.cancel'), { name: "cancel" });
    var okBtn = buttonGroup.add("button", undefined, getLocalizedText('button.ok'), { name: "ok" });

    if (dialog.show() !== 1) return;

    // 選択されたソースとオプションを確定 / Resolve the chosen source and options
    var importFromFolder = folderRadio.value;
    var importByArtboard = byArtboardCheckbox.value;
    var includeGuides = includeGuidesCheckbox.value;
    var originalDocs = importFromFolder ? folderFiles : openDocs;
    if (originalDocs.length < 1) {
        alert(getLocalizedText('alert.noValidFile'));
        return;
    }
    var originalDocsLength = originalDocs.length;

    // スケール（％）。チェックOFFや無効値は 100%（等倍）/ Scale percent; 100% when unchecked or invalid
    var scalePercent = 100;
    if (scaleCheckbox.value) {
        var scaleValue = parseFloat(scaleInput.text);
        if (!isNaN(scaleValue) && scaleValue > 0) scalePercent = scaleValue;
    }

    // 数値はプログレス表示の前に検証する（無効ならパレットを残さず終了）
    // Validate the numbers before showing the progress palette (so it isn't left open on error)
    var docWidthValue = parseFloat(widthInput.text);
    var docHeightValue = parseFloat(heightInput.text);
    if (isNaN(docWidthValue) || isNaN(docHeightValue)) {
        alert(getLocalizedText('alert.invalidNumber'));
        return;
    }

    // フォルダ読み込みで開いた一時ファイルは常に閉じる。開いているファイルのみ「読み込み後の動作」に従う。
    // Temp files opened from a folder are always closed; only open files honor the "After Import" choice.
    var shouldCloseSource = importFromFolder || closeRadio.value;

    // 開いているファイルを閉じる場合、未保存変更があれば確認（変更が失われるため）
    // When closing open files, confirm if any have unsaved changes (they would be lost)
    if (!importFromFolder && closeRadio.value) {
        var hasUnsavedChanges = false;
        for (var d = 0; d < originalDocs.length; d++) {
            if (!originalDocs[d].saved) {
                hasUnsavedChanges = true;
                break;
            }
        }
        if (hasUnsavedChanges && !confirm(getLocalizedText('confirm.discardUnsaved'))) {
            return;
        }
    }

    // プログレスバーのダイアログを表示
    var progressWin = new Window("palette", getLocalizedText('progress.title'));
    progressWin.orientation = "column";
    progressWin.alignChildren = ["fill", "top"];
    progressWin.margins = 20;
    var progressTextGroup = progressWin.add("group");
    progressTextGroup.alignment = ["center", "top"];
    var processedCountStatic = progressTextGroup.add("statictext", undefined, labelText('progress.count') + "0/" + originalDocsLength);
    processedCountStatic.preferredSize = [100, 30];

    var progressBar = progressWin.add("progressbar", undefined, 0, originalDocsLength);
    progressBar.preferredSize = [300, 6];

    var cancelGroup = progressWin.add("group");
    cancelGroup.alignment = "right";
    var progressCancelBtn = cancelGroup.add("button", undefined, getLocalizedText('button.cancel'));

    var userCancelled = false;

    progressCancelBtn.onClick = function () {
        userCancelled = true;
    };
    progressWin.addEventListener("keydown", function (e) {
        if (e.keyName === "Escape") {
            userCancelled = true;
        }
    });
    progressWin.show();

    // 入力値を現在の単位からポイントへ換算 / Convert the input values from the current unit to points
    var ptPerUnit = UNIT_TO_PT[currentUnit];
    var docWidthPt = docWidthValue * ptPerUnit;
    var docHeightPt = docHeightValue * ptPerUnit;
    var colorSpace = rgbRadio.value ? DocumentColorSpace.RGB : DocumentColorSpace.CMYK;

    // 通常ドキュメントの最大寸法（pt）。これを超える場合はラージカンバスとして作成。
    // Max dimension (pt) of a standard document; beyond this, create as a large canvas.
    var STANDARD_MAX_PT = 16383;

    var newDoc;
    if (docWidthPt > STANDARD_MAX_PT || docHeightPt > STANDARD_MAX_PT) {
        var largeCanvasPreset = new DocumentPreset();
        largeCanvasPreset.units = RulerUnits.Points;
        largeCanvasPreset.width = docWidthPt;
        largeCanvasPreset.height = docHeightPt;
        largeCanvasPreset.colorMode = colorSpace;
        largeCanvasPreset.numArtboards = 1;
        newDoc = app.documents.addDocument("Print", largeCanvasPreset);
    } else {
        newDoc = app.documents.add(colorSpace, docWidthPt, docHeightPt);
    }
    app.activeDocument = newDoc;

    // 初期（仮）アートボードの中心をキャンバス中央とみなす / Use the initial (placeholder) artboard's center as the canvas center
    var artboardRect = newDoc.artboards[0].artboardRect; // [L, T, R, B]
    var canvasCenterX = (artboardRect[0] + artboardRect[2]) / 2;
    var canvasCenterY = (artboardRect[1] + artboardRect[3]) / 2;

    // 配置の状態をまとめて保持。各セルを記録し、ループ後に正方形グリッドへ中央配置する。
    // Shared placement state. Each cell is recorded, then laid out into a centered square grid after the loop.
    var placementContext = {
        newDoc: newDoc,
        cells: [],
        canvasCenterX: canvasCenterX,
        canvasCenterY: canvasCenterY,
        byArtboard: importByArtboard,
        placedCount: 0,
        showLabel: showLabelCheckbox.value,
        artboardPadding: CONFIG.artboardMargin,
        scalePercent: scalePercent
    };

    // ループ全体を try/finally で囲み、エラーが出てもプログレスは必ず閉じる
    // Wrap the whole loop so the progress palette is always closed, even on error
    try {
        for (var j = 0; j < originalDocs.length; j++) {
            $.sleep(0);
            app.redraw();
            // キャンセルされたらループを抜けて、ここまでの結果で後始末する / On cancel, break and finish with what was placed so far
            if (userCancelled) break;

            var srcDoc = importFromFolder ? app.open(originalDocs[j]) : originalDocs[j];

            // 1ファイル分の処理。エラーが出ても finally で一時ファイルを必ず閉じ、状態も復元する。
            // Process one file; finally always closes the temp file and restores state, even on error.
            var lockState = null;
            try {
                app.activeDocument = srcDoc;
                var labelName = srcDoc.name.replace(/\.[^\.]+$/, "");

                if (importByArtboard) {
                    // アートボード単位：ロック・非表示も含め、各アートボード上のオブジェクトを取り込む
                    // Per artboard: import every object on each artboard, including locked/hidden ones
                    // 開いているファイルを閉じない場合は、解除前の状態を記録して後で復元する
                    // When keeping an open file, capture the state before unlocking so it can be restored afterward
                    lockState = shouldCloseSource ? null : captureLockHiddenState(srcDoc);
                    unlockAllLayersAndItems(srcDoc);
                    for (var ab = 0; ab < srcDoc.artboards.length; ab++) {
                        srcDoc.artboards.setActiveArtboardIndex(ab);
                        srcDoc.selection = null;
                        srcDoc.selectObjectsOnActiveArtboard();
                        if (srcDoc.selection.length === 0) continue;

                        var abRect = srcDoc.artboards[ab].artboardRect; // [L, T, R, B]

                        // ガイドを含める場合は、このアートボードに重なる短いガイドを選択へ追加（相対位置の算出前に）
                        // When including guides, add the short guides overlapping this artboard before measuring bounds
                        if (includeGuides) {
                            var abGuides = collectImportableGuides(srcDoc, abRect[2] - abRect[0], abRect[1] - abRect[3], abRect);
                            if (abGuides.length > 0) {
                                var combined = [];
                                for (var g = 0; g < srcDoc.selection.length; g++) combined.push(srcDoc.selection[g]);
                                srcDoc.selection = combined.concat(abGuides);
                            }
                        }

                        // 元のアートボードサイズと、その中での内容（＋ガイド）の相対位置を記録
                        // Record the original artboard size and the relative position of the content (and guides)
                        var selBounds = getCombinedVisibleBounds(srcDoc.selection);
                        var artboardCell = {
                            width: abRect[2] - abRect[0],
                            height: abRect[1] - abRect[3],
                            offsetX: selBounds[0] - abRect[0],
                            offsetY: abRect[1] - selBounds[1]
                        };

                        app.copy();
                        app.activeDocument = newDoc;
                        app.paste();
                        var placedArtboard = placePastedGroup(placementContext, labelName, artboardCell);
                        app.activeDocument = srcDoc;
                        // 配置に失敗したアートボードはスキップ（既にアラート済み）/ Skip artboards that failed to place (already alerted)
                        if (!placedArtboard) continue;
                    }
                } else {
                    // 通常：現在選択可能な表示中オブジェクトのみ取り込む
                    // Default: import only currently selectable, visible objects
                    app.executeMenuCommand("selectall");

                    var filteredSelection = [];
                    for (var s = 0; s < srcDoc.selection.length; s++) {
                        var obj = srcDoc.selection[s];
                        if (!obj.locked && !obj.hidden) {
                            filteredSelection.push(obj);
                        }
                    }
                    // ガイドを含める場合は、短いガイド（カンバス＝先頭アートボード基準）を追加
                    // When including guides, add the short guides (canvas = the first artboard)
                    if (includeGuides) {
                        var firstArtboard = srcDoc.artboards[0].artboardRect;
                        filteredSelection = filteredSelection.concat(
                            collectImportableGuides(srcDoc, firstArtboard[2] - firstArtboard[0], firstArtboard[1] - firstArtboard[3], null)
                        );
                    }
                    if (filteredSelection.length > 0) {
                        srcDoc.selection = filteredSelection;
                        app.copy();
                        app.activeDocument = newDoc;
                        app.paste();
                        placePastedGroup(placementContext, labelName);
                    }
                }
            } finally {
                if (lockState) restoreLockHiddenState(lockState);                 // 元の状態へ戻す / Restore the original state
                if (shouldCloseSource) srcDoc.close(SaveOptions.DONOTSAVECHANGES); // 一時ファイルを閉じる / Close the temp file
            }

            progressBar.value = j + 1;
            processedCountStatic.text = labelText('progress.count') + (j + 1) + "/" + originalDocsLength;
            progressWin.update();
        }

        // 全セルを正方形に近いグリッドへ並べ、キャンバス中央に配置
        // Lay out all cells into a near-square grid centered on the canvas
        layoutCellsAsCenteredGrid(placementContext);

        // グループごとにアートボードを追加したので、初期の仮アートボードを削除
        // Each group added its own artboard, so remove the initial placeholder artboard
        if (placementContext.placedCount > 0 && newDoc.artboards.length > placementContext.placedCount) {
            newDoc.artboards.remove(0);
        }
        app.activeDocument = newDoc;
    } finally {
        progressWin.close();
    }

    // すべてのアートボードがウィンドウに収まるように表示 / Fit all artboards in the window
    app.executeMenuCommand('fitall');

    // キャンセルされた場合はエラーではなく通常のメッセージで知らせる
    // If cancelled, inform the user with a normal message (not an error)
    if (userCancelled) {
        alert(getLocalizedText('alert.cancelled'));
    }
})();