#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

SmartBatchImporter.jsx

### 概要

- 複数のIllustratorファイル（.ai/.svg/.eps）を一括で読み込み、取り込んだファイル（またはアートボード）ごとに1つのアートボードを作成して、全体が正方形に近くなるグリッドへ整列配置するスクリプトです。
- 読み込み対象の選択、ファイル形式／ファイル名フィルター、読み込み先（新規ドキュメント／現在のドキュメント）、対象アートボードの指定、アートボード単位読み込み、ガイドの取り込み、スケール、ファイル名ラベルなど多彩なオプションを備えています。

### 主な機能

- 開いているファイル、またはフォルダー指定による一括読み込み（ファイル形式：AI / SVG / EPS）
- フィルター：ファイル形式の絞り込みと、ファイル名の正規表現による絞り込み
- 読み込み先を選択：新規ドキュメントを作成、または現在のドキュメントへ取り込み（現在のドキュメント時は既存アートボードの下に重ならないよう配置）
- 新規ドキュメント設定：カラーモード（RGB/CMYK）、ラスタライズ効果の解像度（72/150/300 ppi）、サイズプリセット（カスタム／A4／フルHD／ラージカンバス、単位 mm/px・A4 は CMYK 既定）
- 対象アートボードの指定：1のみ／すべて／指定（例: 1, 3-5）
- 取り込んだファイル（またはアートボード）ごとに1つのアートボードを作成し、全体が正方形に近くなるよう列数・行数を自動調整して整列（完了後に全体を表示）
- アートボード単位での読み込み（対象アートボードのロック・非表示も対象、元のアートボードサイズと相対位置を保持）
- ガイドの取り込み（ルーラーガイド＝カンバスの半分以上に伸びる長いガイドは除外）
- スケール（％）指定（アートボード枠・内容・線幅を同率で）
- ファイル名ラベルの追加（フォルダー指定時は既定OFF）
- 読み込み後にファイルを閉じる／保持（開いているファイル選択時のみ。未保存変更があれば確認）
- プログレスバーとキャンセル対応
- 日本語／英語インターフェース対応

### 紹介記事

https://note.com/dtp_tranist/n/n8180588e5630

### 更新履歴

- v1.0.0 (20250529) : 初期バージョン
- v1.2.0 (20260613) : 読み込み対象選択（開いているファイル／フォルダー指定＋種別フィルタ AI/SVG）、新しいドキュメント設定（カラースペース・サイズプリセット・単位自動切替・ラージカンバス）、アートボード単位読み込み（ロック・非表示対象／元サイズ保持）、スケール、ファイル名ラベル、読み込み後の動作
- v1.2.1 (20260613) : 種別フィルタに EPS を追加、ダイアログのパネル構成を2カラム化（読み込み対象／種別、カラースペース／ドキュメントサイズ、読み込みオプション）
- v1.3.0 (20260613) : 読み込み先に「現在のドキュメント」を追加（既存アートボードの下に重ならず配置／自分自身はソースから除外）、対象アートボード指定（1のみ／すべて／指定 例 1,3-5）、ファイル名の正規表現フィルター、ラスタライズ効果の解像度指定（72/150/300）、サイズの単位を mm/px に整理（A4 は CMYK 既定）、「種別→ファイル形式」「カラースペース→カラーモード」等の名称整理、変数・パネル・関数名を整理

---

### Script Name:

SmartBatchImporter.jsx

### Overview

- A script to batch import multiple Illustrator files (.ai/.svg/.eps), turn each imported file (or artboard) into one artboard, and arrange them into a near-square grid.
- Offers source selection, file-format / file-name filtering, a destination choice (new or current document), target-artboard selection, per-artboard import, guide import, scaling, filename labels, and more.

### Main Features

- Batch import from currently open files or a chosen folder (formats: AI / SVG / EPS)
- Filter: narrow by file format and by a file-name regular expression
- Choose the destination: create a new document, or import into the current document (placed below the existing artboards without overlapping)
- New-document settings: color mode (RGB/CMYK), raster effects resolution (72/150/300 ppi), size presets (Custom / A4 / Full HD / Large Canvas; mm/px units; A4 defaults to CMYK)
- Target artboards: only the first, all, or a specified set (e.g. 1, 3-5)
- One artboard per imported file (or artboard); auto-adjusts columns/rows toward a near-square layout (fits all in the window when done)
- Per-artboard import (includes locked/hidden objects on the target artboards, preserves the original artboard size and relative positions)
- Imports guides (excludes ruler guides — long guides that span half the canvas or more)
- Scale (%) option (artboard frame, content, and stroke widths at the same rate)
- Adds file-name labels (off by default in folder mode)
- Close or keep source files after import (only when importing open files; confirms if there are unsaved changes)
- Progress bar with cancel support
- Japanese and English UI support

### Update History

- v1.0.0 (20250529): Initial version
- v1.2.0 (20260613): Source selection (open files / folder + AI/SVG type filter), new-document settings (color space, size presets, auto unit switching, Large Canvas), per-artboard import (incl. locked/hidden, preserves original size), scale, file-name label, and after-import action
- v1.2.1 (20260613): Added EPS to the type filter; reorganized the dialog into two-column panels (source / type, color space / document size, import options)
- v1.3.0 (20260613): Added "Current document" as a destination (placed below existing artboards without overlap; the target itself is excluded from sources), target-artboard selection (only 1 / all / specify, e.g. 1,3-5), a file-name regular-expression filter, raster effects resolution (72/150/300), mm/px size units (A4 defaults to CMYK), label cleanups ("Type→File format", "Color Space→Color Mode", etc.), and tidied variable/panel/function names
*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.3.0";

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
        filter: { ja: "フィルター", en: "Filter" },
        destination: { ja: "読み込み先", en: "Destination" },
        colorMode: { ja: "カラーモード", en: "Color Mode" },
        resolution: { ja: "解像度", en: "Resolution" },
        docSize: { ja: "サイズ", en: "Size" },
        options: { ja: "読み込みオプション", en: "Import Options" },
        afterImport: { ja: "読み込み後", en: "After Import" }
    },
    radio: {
        openFiles: { ja: "現在、開いているファイル", en: "Currently open files" },
        specifyFolder: { ja: "フォルダーを指定", en: "Specify folder" },
        rgb: { ja: "RGB", en: "RGB" },
        cmyk: { ja: "CMYK", en: "CMYK" },
        currentDoc: { ja: "現在のドキュメント", en: "Current document" },
        newDoc: { ja: "新規ドキュメント", en: "New document" },
        artboardOne: { ja: "1のみ", en: "Artboard 1 only" },
        artboardAll: { ja: "すべて", en: "All" },
        artboardSpecify: { ja: "指定", en: "Specify" },
        closeDoc: { ja: "閉じる", en: "Close" },
        keepOpen: { ja: "開いたまま", en: "Keep Open" }
    },
    checkbox: {
        byArtboard: { ja: "アートボード単位", en: "Import per artboard" },
        includeGuides: { ja: "ガイド", en: "Guides" },
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
        fileType: { ja: "ファイル形式", en: "File format" },
        fileName: { ja: "ファイル名", en: "File name" },
        unit: { ja: "単位", en: "Unit" },
        artboardTarget: { ja: "対象アートボード", en: "Target artboards" }
    },
    hint: {
        filter: {
            ja: "正規表現を使ってファイルを絞り込み",
            en: "Filter files using a regular expression"
        }
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
        noCurrentDoc: {
            ja: "「現在のドキュメント」に読み込むには、ドキュメントを開いておいてください。",
            en: "Open a document first to import into the current document."
        },
        invalidArtboardSpec: {
            ja: "対象アートボードの指定が正しくありません。番号で指定してください。例: 1, 3-5",
            en: "The target artboard specification is invalid. Specify by number, e.g. 1, 3-5"
        },
        noArtboardImported: {
            ja: "取り込める内容が見つかりませんでした。対象アートボードの番号やファイルの内容を確認してください。",
            en: "Nothing could be imported. Check the target artboard numbers and the file contents."
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
            ja: "読み込みを中断しました。ここまでに読み込んだ内容はドキュメントに残っています。",
            en: "Import was stopped. Items imported so far remain in the document."
        },
        invalidScale: {
            ja: "拡大・縮小の値が正しくありません。0より大きい数値を入力してください。",
            en: "The scale value is invalid. Enter a number greater than 0."
        },
        invalidFilter: {
            ja: "ファイル名フィルターの正規表現が正しくありません。式を見直すか、空欄にしてください。",
            en: "The file-name filter is not a valid regular expression. Fix it or clear the field."
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
            ja: "指定したフォルダー内の .ai / .svg / .eps ファイルを読み込み対象にします。",
            en: "Uses .ai / .svg / .eps files in the chosen folder as the import source."
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
            ja: "読み込むファイル（.ai / .svg / .eps）が入ったフォルダーを選びます。",
            en: "Choose a folder that contains the files to import (.ai / .svg / .eps)."
        },
        currentDoc: {
            ja: "現在開いているドキュメントに読み込みます。下のカラーモード・解像度・サイズ設定は使いません。",
            en: "Imports into the currently active document. The color mode / resolution / size settings below are not used."
        },
        newDoc: {
            ja: "下の設定で新規ドキュメントを作成し、そこに読み込みます。",
            en: "Creates a new document with the settings below and imports into it."
        },
        colorMode: {
            ja: "新規ドキュメントのカラーモードを選びます。読み込み元の色を完全に変換する機能ではありません。",
            en: "Choose the color mode for the new document. This does not fully convert all colors from source files."
        },
        byArtboard: {
            ja: "各アートボードを別々に読み込み、元のアートボードサイズと相対位置を保ちます。ロック・非表示オブジェクトも対象です。",
            en: "Imports each artboard separately and preserves the original artboard size and relative positions. Locked/hidden objects are included."
        },
        artboardTarget: {
            ja: "取り込むアートボードを選びます。「1のみ」「すべて」「指定」から選択します。",
            en: "Choose which artboards to import: only the first, all, or a specified set."
        },
        artboardSpecify: {
            ja: "取り込むアートボードを番号で指定します。例: 1, 3-5（カンマ区切り・範囲指定可）。",
            en: "Specify artboards to import by number, e.g. 1, 3-5 (comma-separated, ranges allowed)."
        },
        includeGuides: {
            ja: "ルーラーガイド（カンバスの半分以上に伸びる長いガイド）を除いて、ガイドを読み込みます。",
            en: "Imports guides, excluding ruler guides (long guides that span half the canvas or more)."
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
    inch: 72      // 1inch = 72pt（ラージカンバスのネイティブ単位、換算用）
};

// 単位ドロップダウンの並び / Units shown in the unit dropdown
var UNIT_LIST = ["mm", "px"];

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
            var layer = layerList[i];
            layers.push({ ref: layer, locked: layer.locked, visible: layer.visible });
            walkLayers(layer.layers);
        }
    }
    walkLayers(doc.layers);
    for (var pi = 0; pi < doc.pageItems.length; pi++) {
        var pageItem = doc.pageItems[pi];
        items.push({ ref: pageItem, locked: pageItem.locked, hidden: pageItem.hidden });
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
function rectsIntersect(rectA, rectB) {
    return !(rectA[2] < rectB[0] || rectA[0] > rectB[2] || rectA[3] > rectB[1] || rectA[1] < rectB[3]);
}

/* 指定アートボード（複数可）に重なるオブジェクトだけロック・非表示を解除する（他アートボードのオブジェクトには触れない）。
   レイヤーは選択可能にするため一時的に全解除する。戻り値は restoreLockHiddenState で復元できる状態。
   / Unlock only the objects overlapping the given artboard rects (leaving other artboards' objects untouched).
   Layers are unlocked/shown temporarily so the items become selectable. The result restores via restoreLockHiddenState. */
function unlockItemsOnArtboards(doc, abRects) {
    var layers = [];
    function walkLayers(layerList) {
        for (var i = 0; i < layerList.length; i++) {
            var layer = layerList[i];
            layers.push({ ref: layer, locked: layer.locked, visible: layer.visible });
            layer.locked = false;
            layer.visible = true;
            walkLayers(layer.layers);
        }
    }
    walkLayers(doc.layers);

    var items = [];
    for (var pi = 0; pi < doc.pageItems.length; pi++) {
        var pageItem = doc.pageItems[pi];
        if (!pageItem.locked && !pageItem.hidden) continue; // 既に選択可能なものは触らない / Leave already-selectable items alone
        var bounds;
        try { bounds = pageItem.geometricBounds; } catch (eBounds) { continue; }
        var overlapsTarget = false;
        for (var ri = 0; ri < abRects.length; ri++) {
            if (rectsIntersect(bounds, abRects[ri])) { overlapsTarget = true; break; }
        }
        if (!overlapsTarget) continue; // 対象アートボードのどれにも重ならないものは対象外 / Skip items outside all target artboards
        items.push({ ref: pageItem, locked: pageItem.locked, hidden: pageItem.hidden });
        pageItem.locked = false;
        pageItem.hidden = false;
    }
    return { layers: layers, items: items };
}

/* "1, 3-5" のような指定文字列を 1 始まりのアートボード番号の配列にする（カンマ区切り・範囲対応）。
   / Parse a spec like "1, 3-5" into an array of 1-based artboard numbers (comma-separated, ranges supported). */
function parseArtboardNumbers(spec) {
    var numbers = [];
    if (!spec) return numbers;
    var parts = spec.split(/[,，\s]+/);
    for (var i = 0; i < parts.length; i++) {
        var part = parts[i];
        if (part === "") continue;
        var range = part.match(/^(\d+)\s*[-–~]\s*(\d+)$/);
        if (range) {
            var from = parseInt(range[1], 10);
            var to = parseInt(range[2], 10);
            if (from > to) { var swap = from; from = to; to = swap; }
            for (var n = from; n <= to; n++) numbers.push(n);
        } else if (/^\d+$/.test(part)) {
            numbers.push(parseInt(part, 10));
        }
    }
    return numbers;
}

/* 対象モードと指定文字列から、取り込むアートボードの 0 始まりインデックス配列を返す。
   / Resolve the 0-based artboard indices to import from the mode and spec text. */
function resolveTargetArtboardIndices(doc, mode, specText) {
    var count = doc.artboards.length;
    var indices = [];
    if (mode === "all") {
        for (var i = 0; i < count; i++) indices.push(i);
    } else if (mode === "specify") {
        var numbers = parseArtboardNumbers(specText);
        var seen = {};
        for (var k = 0; k < numbers.length; k++) {
            var idx = numbers[k] - 1; // 1 始まり → 0 始まり / 1-based → 0-based
            if (idx >= 0 && idx < count && !seen[idx]) { seen[idx] = true; indices.push(idx); }
        }
        indices.sort(function (a, b) { return a - b; });
    } else { // "first"
        if (count > 0) indices.push(0);
    }
    return indices;
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
        var guideBounds = item.geometricBounds; // [L, T, R, B]
        if ((guideBounds[2] - guideBounds[0]) >= halfWidth || (guideBounds[1] - guideBounds[3]) >= halfHeight) continue; // 長すぎるガイドは無視
        if (withinRect && !rectsIntersect(guideBounds, withinRect)) continue;
        result.push(item);
    }
    return result;
}

/* ドキュメント内の全アートボードを囲う合成矩形 [L, T, R, B] を返す / Union rect [L, T, R, B] of all artboards in the doc */
function getArtboardsUnionRect(doc) {
    var firstRect = doc.artboards[0].artboardRect; // [L, T, R, B]
    var left = firstRect[0], top = firstRect[1], right = firstRect[2], bottom = firstRect[3];
    for (var i = 1; i < doc.artboards.length; i++) {
        var rect = doc.artboards[i].artboardRect;
        if (rect[0] < left) left = rect[0];
        if (rect[1] > top) top = rect[1];
        if (rect[2] > right) right = rect[2];
        if (rect[3] < bottom) bottom = rect[3];
    }
    return [left, top, right, bottom];
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

/* 貼り付け直後のオブジェクトを1グループにまとめてセルとして記録し、必要ならラベルを付ける。
   ctx.createArtboards が true のときだけセルに合わせたアートボードを追加する（OFF時は枠を作らず内容だけ）。
   / Group the just-pasted objects into one cell and optionally add a label.
   Adds a matching artboard only when ctx.createArtboards is true (otherwise the content is placed without a frame).
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
    var cellPadding = ctx.artboardPadding;
    var cellWidth, cellHeight, offsetX, offsetY;
    if (artboardCell) {
        // 元のアートボードサイズと相対位置を保持（スケールも反映）/ Preserve the original artboard size and relative position (scaled)
        cellWidth = artboardCell.width * scaleFactor;
        cellHeight = artboardCell.height * scaleFactor;
        offsetX = artboardCell.offsetX * scaleFactor;
        offsetY = artboardCell.offsetY * scaleFactor;
    } else {
        // 内容に余白を足した枠 / A box fitted to the content plus padding
        cellWidth = contentWidth + cellPadding * 2;
        cellHeight = contentHeight + cellPadding * 2;
        offsetX = cellPadding;
        offsetY = cellPadding;
    }

    // 内容の現在位置を基準にしたセルの左上（最終位置は後でグリッド配置）
    // Cell's top-left based on the content's current spot (final position is set later by the grid layout)
    var abLeft = bounds[0] - offsetX;
    var abTop = bounds[1] + offsetY;

    // アートボードを作るのは「アートボード単位」ONのときだけ。OFFのときは枠を作らず内容だけ配置する。
    // Create an artboard only when "per artboard" is on; when off, place the content without adding a frame.
    var newArtboard = null;
    if (ctx.createArtboards) {
        newArtboard = newDoc.artboards.add([abLeft, abTop, abLeft + cellWidth, abTop - cellHeight]);
        newArtboard.name = labelName;
    }

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

    // セルを記録（グループ・アートボード・ラベルを最終グリッド配置で一緒に動かす）。
    // アートボードを作らない場合のために、現在のセル左上（currentLeft/currentTop）も保持する。
    // Record the cell so the group, artboard, and label move together in the final grid layout.
    // Also keep the current cell top-left (currentLeft/currentTop) for the no-artboard case.
    ctx.cells.push({
        group: pastedGroup,
        artboard: newArtboard, // OFF時は null / null when no artboard is created
        label: labelItem,
        width: cellWidth,
        height: cellHeight,
        currentLeft: abLeft,
        currentTop: abTop
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

    var startLeft, startTop;
    if (ctx.avoidRect) {
        // 現在のドキュメント：既存アートボードの下に、重ならないよう配置（水平は既存の中央に揃える）
        // Current document: place below the existing artboards (no overlap), centered horizontally on them
        var existingRect = ctx.avoidRect; // [L, T, R, B]
        var existingCenterX = (existingRect[0] + existingRect[2]) / 2;
        startLeft = existingCenterX - totalWidth / 2;
        startTop = existingRect[3] - gapY; // 既存アートボードの下端からひと間隔あけた位置をグリッド上端に / Grid top sits one gap below the existing bottom
    } else {
        // グリッド全体をキャンバス中央に揃える / Center the whole grid on the canvas
        startLeft = ctx.canvasCenterX - totalWidth / 2;
        startTop = ctx.canvasCenterY + totalHeight / 2;
    }

    for (var k = 0; k < count; k++) {
        var col = k % columns;
        var row = Math.floor(k / columns);
        var slotLeft = startLeft + col * (slotWidth + gapX);
        var slotTop = startTop - row * (slotHeight + gapY);

        var cell = cells[k];
        // セルをスロット内で中央に / Center the cell within its slot
        var targetLeft = slotLeft + (slotWidth - cell.width) / 2;
        var targetTop = slotTop - (slotHeight - cell.height) / 2;

        // アートボードがある場合はその枠、無い場合は記録した現在のセル左上を基準に移動量を求める
        // Use the artboard frame if present; otherwise the recorded current cell top-left
        var rect = cell.artboard
            ? cell.artboard.artboardRect // [L, T, R, B]
            : [cell.currentLeft, cell.currentTop, cell.currentLeft + cell.width, cell.currentTop - cell.height];
        var dx = targetLeft - rect[0];
        var dy = targetTop - rect[1];

        cell.group.translate(dx, dy);
        if (cell.artboard) cell.artboard.artboardRect = [rect[0] + dx, rect[1] + dy, rect[2] + dx, rect[3] + dy];
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
    var folderFiles = [];        // 種別＋名前フィルタ後のファイル / Files after the type + name filter
    var folderFilesTotal = 0;    // 種別フィルタに一致する総数（名前フィルタ前）/ Total matching the type filter (before the name filter)

    var dialog = new Window("dialog", getLocalizedText('dialog.title') + ' ' + SCRIPT_VERSION);
    dialog.orientation = "column";
    dialog.alignChildren = ["left", "top"];

    // --- 読み込み対象パネル / Source panel ---
    var sourcePanel = dialog.add("panel", undefined, getLocalizedText('panel.source'));
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
    folderNameText.preferredSize = [360, 20];
    folderNameText.minimumSize = [360, 20];

    // --- フィルターパネル（読み込み対象パネル内。種別＋ファイル名の正規表現で絞り込み）/ Filter panel (nested in the source panel) ---
    var filterPanel = sourcePanel.add("panel", undefined, getLocalizedText('panel.filter'));
    setupPanel(filterPanel);

    // 種別（読み込むファイル形式、チェックボックスは横並び）/ File types (checkboxes laid out horizontally)
    var typeRow = filterPanel.add("group");
    setupGroup(typeRow, "row");
    var typeLabel = typeRow.add("statictext", undefined, labelText('field.fileType'));
    typeLabel.preferredSize = [100, 20];
    var aiCheckbox = typeRow.add("checkbox", undefined, "AI");
    aiCheckbox.helpTip = getLocalizedText('tooltip.fileType');
    var svgCheckbox = typeRow.add("checkbox", undefined, "SVG");
    svgCheckbox.helpTip = getLocalizedText('tooltip.fileType');
    var epsCheckbox = typeRow.add("checkbox", undefined, "EPS");
    epsCheckbox.helpTip = getLocalizedText('tooltip.fileType');
    aiCheckbox.value = true;   // 既定は AI のみ ON / Default: AI only
    svgCheckbox.value = false;
    epsCheckbox.value = false;

    var filterRow = filterPanel.add("group");
    setupGroup(filterRow, "row");
    var filterLabel = filterRow.add("statictext", undefined, labelText('field.fileName'));
    filterLabel.preferredSize = [100, 20];
    var filterInput = filterRow.add("edittext", undefined, "");
    filterInput.characters = 20;
    filterInput.helpTip = getLocalizedText('hint.filter');
    filterLabel.helpTip = getLocalizedText('hint.filter');

    // ファイル数はパネルのタイトルに表示する。フィルター使用時は「読み込み対象（3/5）」（絞り込み後/総数）、
    // 未使用時は「読み込み対象（5）」のように表示する。
    // Show the file count in the panel title. With an active filter, "Source (3/5)" (filtered/total); otherwise "Source (5)".
    function updateFileCount() {
        var filterActive = (filterInput.text !== "" && getNameFilterRegExp() !== null);
        var total, filtered;
        if (folderRadio.value) {
            total = folderFilesTotal;        // 種別フィルタに一致する総数 / total matching the type filter
            filtered = folderFiles.length;   // 種別＋名前フィルタ後 / after the type + name filter
        } else {
            total = openDocs.length;
            filtered = getFilteredOpenDocs().length;
        }
        var countText = filterActive ? (filtered + '/' + total) : String(total);
        sourcePanel.text = getLocalizedText('panel.source') + (currentLanguage === 'ja' ? '（' + countText + '）' : ' (' + countText + ')');
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
        typeLabel.enabled = folderRadio.value;
        aiCheckbox.enabled = folderRadio.value;
        svgCheckbox.enabled = folderRadio.value;
        epsCheckbox.enabled = folderRadio.value;
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

    // ファイル名フィルタの正規表現を返す。空欄や不正な式なら null（＝全件対象）。
    // Return the file-name filter as a RegExp; null when empty or invalid (= match all).
    function getNameFilterRegExp() {
        var pattern = filterInput.text;
        if (pattern === null || pattern === "") return null;
        try {
            return new RegExp(pattern, "i");
        } catch (e) {
            return null; // 入力途中の不正な式では絞り込まない / Don't filter on an incomplete/invalid pattern
        }
    }

    // ファイル名がフィルタに一致するか / Whether a file name matches the current filter
    function matchesNameFilter(name) {
        var nameRe = getNameFilterRegExp();
        return nameRe === null || nameRe.test(name);
    }

    // 開いているファイルをフィルタで絞り込んだ配列 / Open documents narrowed by the name filter
    function getFilteredOpenDocs() {
        var matched = [];
        for (var i = 0; i < openDocs.length; i++) {
            if (matchesNameFilter(openDocs[i].name)) matched.push(openDocs[i]);
        }
        return matched;
    }

    // 選択中フォルダを現在の種別とファイル名フィルタで再フィルタ / Re-filter the chosen folder by file types and the name filter
    // 種別に一致する総数（folderFilesTotal）と、名前フィルタ後の一覧（folderFiles）の両方を更新する。
    // Updates both the type-matched total (folderFilesTotal) and the name-filtered list (folderFiles).
    function refreshFolderFiles() {
        folderFiles = [];
        folderFilesTotal = 0;
        var exts = getSelectedExtensions();
        if (selectedFolder && exts.length > 0) {
            var typeRe = new RegExp("\\.(" + exts.join("|") + ")$", "i");
            var typeMatched = selectedFolder.getFiles(function (f) {
                return f instanceof File && typeRe.test(f.name);
            });
            folderFilesTotal = typeMatched.length;
            var matched = [];
            for (var i = 0; i < typeMatched.length; i++) {
                if (matchesNameFilter(typeMatched[i].name)) matched.push(typeMatched[i]);
            }
            matched.sort(function (a, b) {
                return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1;
            });
            folderFiles = matched;
        }
        updateFileCount();
    }

    // フィルタ入力に応じて件数・一覧を更新 / Refresh counts and lists as the filter changes
    filterInput.onChanging = function () {
        refreshFolderFiles(); // updateFileCount を内包（開いているファイル件数も更新）/ also refreshes the open-files count
    };

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

    // --- 読み込み先パネル（読み込み先の選択＋新規ドキュメント設定）/ Destination panel (target choice + new-document settings) ---
    var destinationPanel = dialog.add("panel", undefined, getLocalizedText('panel.destination'));
    setupPanel(destinationPanel);

    // 読み込み先：現在のドキュメント／新規ドキュメント / Destination: current document or a new one
    var destRow = destinationPanel.add("group");
    setupGroup(destRow, "row");
    destRow.alignment = ["center", "top"]; // 左右中央 / Center horizontally
    destRow.margins = [0, 0, 0, 10];        // 下に10pxの余白 / 10px margin below
    var currentDocRadio = destRow.add("radiobutton", undefined, getLocalizedText('radio.currentDoc'));
    currentDocRadio.helpTip = getLocalizedText('tooltip.currentDoc');
    var newDocRadio = destRow.add("radiobutton", undefined, getLocalizedText('radio.newDoc'));
    newDocRadio.helpTip = getLocalizedText('tooltip.newDoc');
    newDocRadio.value = true; // 既定は新規ドキュメント / Default: new document

    // カラーモード＋解像度とサイズを横2カラムで並べる / Lay out color mode + resolution and size in two columns
    var newDocSettingsRow = destinationPanel.add("group");
    setupGroup(newDocSettingsRow, "row");
    newDocSettingsRow.alignChildren = ["left", "fill"]; // 2つのパネルの高さを揃える / Match the two panels' heights
    newDocSettingsRow.spacing = 15;                     // カラーモード／解像度とサイズの2カラム間隔 / Gap between the color-mode/resolution and size columns

    // 読み込み先が「現在のドキュメント」のときは新規ドキュメント設定をディムにする
    // Dim the new-document settings when the destination is the current document
    function updateDestinationState() {
        newDocSettingsRow.enabled = newDocRadio.value;
    }
    currentDocRadio.onClick = updateDestinationState;
    newDocRadio.onClick = updateDestinationState;

    // 左カラム：カラーモードと解像度を縦に積む / Left column: color mode + resolution stacked
    var colorAndResolutionColumn = newDocSettingsRow.add("group");
    setupGroup(colorAndResolutionColumn, "column");
    colorAndResolutionColumn.alignChildren = ["fill", "top"];

    // カラーモード（ラジオは縦並び）/ Color mode (radios stacked vertically)
    var colorModePanel = colorAndResolutionColumn.add("panel", undefined, getLocalizedText('panel.colorMode'));
    setupPanel(colorModePanel);
    var rgbRadio = colorModePanel.add("radiobutton", undefined, getLocalizedText('radio.rgb'));
    rgbRadio.helpTip = getLocalizedText('tooltip.colorMode');
    var cmykRadio = colorModePanel.add("radiobutton", undefined, getLocalizedText('radio.cmyk'));
    cmykRadio.helpTip = getLocalizedText('tooltip.colorMode');
    rgbRadio.value = true;

    // 解像度（ラスタライズ効果設定の ppi）/ Resolution (raster effects ppi)
    var resolutionPanel = colorAndResolutionColumn.add("panel", undefined, getLocalizedText('panel.resolution'));
    setupPanel(resolutionPanel);
    var resolutionDropdown = resolutionPanel.add("dropdownlist", undefined, ["72", "150", "300"]);
    resolutionDropdown.selection = 2; // デフォルトは300 / Default 300

    var sizePanel = newDocSettingsRow.add("panel", undefined, getLocalizedText('panel.docSize'));
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
    widthInput.characters = 5;
    widthInput.helpTip = getLocalizedText('tooltip.size');

    var heightRow = sizePanel.add("group");
    setupGroup(heightRow, "row");
    var heightLabel = heightRow.add("statictext", undefined, labelText('field.height'));
    heightLabel.preferredSize = [40, 20];
    var heightInput = heightRow.add("edittext", undefined, "1000");
    heightInput.characters = 5;
    heightInput.helpTip = getLocalizedText('tooltip.size');

    // 単位（mm / px）。A4 は mm、それ以外は px を既定にし、手動切替で値を換算する。
    // Unit (mm / px). Default mm for A4, px otherwise; switching converts the values.
    var unitRow = sizePanel.add("group");
    setupGroup(unitRow, "row");
    var unitLabel = unitRow.add("statictext", undefined, labelText('field.unit'));
    unitLabel.preferredSize = [40, 20];
    var unitDropdown = unitRow.add("dropdownlist", undefined, UNIT_LIST);

    // 入力値を旧単位から新単位へ換算する（物理サイズを保つ）/ Convert an input value from the old unit to the new one (keeps physical size)
    function convertUnitText(textValue, fromUnit, toUnit) {
        var value = parseFloat(textValue);
        if (isNaN(value)) return textValue;
        return String(Math.round(value * UNIT_TO_PT[fromUnit] / UNIT_TO_PT[toUnit]));
    }

    // ドロップダウンの選択をプログラムから変更（onChange を誤発火させない）/ Set the dropdown selection without firing onChange
    var suppressUnitChange = false;
    function selectUnit(unitName) {
        suppressUnitChange = true;
        for (var i = 0; i < UNIT_LIST.length; i++) {
            if (UNIT_LIST[i] === unitName) { unitDropdown.selection = i; break; }
        }
        suppressUnitChange = false;
    }
    selectUnit(currentUnit); // 初期は px（カスタム）/ Initial: px (Custom)

    // 手動で単位を切り替えたら現在の値を換算 / Convert current values when the unit is switched manually
    unitDropdown.onChange = function () {
        if (suppressUnitChange || !unitDropdown.selection) return;
        var newUnit = unitDropdown.selection.text;
        if (newUnit === currentUnit) return;
        widthInput.text = convertUnitText(widthInput.text, currentUnit, newUnit);
        heightInput.text = convertUnitText(heightInput.text, currentUnit, newUnit);
        currentUnit = newUnit;
    };

    presetDropdown.onChange = function () {
        var idx = presetDropdown.selection.index;
        var preset = SIZE_PRESETS[idx];
        var isA4 = idx === 1;
        var newUnit = isA4 ? "mm" : "px"; // A4 は mm、それ以外は px / mm for A4, px otherwise
        if (preset.width !== undefined) {
            // プリセット：ネイティブ単位の値を表示単位へ換算 / Preset: convert its native-unit values to the display unit
            widthInput.text = convertUnitText(String(preset.width), preset.unit, newUnit);
            heightInput.text = convertUnitText(String(preset.height), preset.unit, newUnit);
        } else {
            // カスタム：現在の値を新しい単位へ換算して物理サイズを保つ / Custom: convert current values to keep the physical size
            widthInput.text = convertUnitText(widthInput.text, currentUnit, newUnit);
            heightInput.text = convertUnitText(heightInput.text, currentUnit, newUnit);
        }
        currentUnit = newUnit;
        selectUnit(currentUnit);

        // A4（印刷向け）は CMYK、それ以外（画面向け）は RGB を既定にする
        // Default to CMYK for A4 (print), RGB for the others (screen)
        cmykRadio.value = isA4;
        rgbRadio.value = !isA4;
    };

    // --- 読み込みオプションパネル / Import options panel ---
    var optionsPanel = dialog.add("panel", undefined, getLocalizedText('panel.options'));
    setupPanel(optionsPanel);
    // オプションを2カラムで並べる（左：アートボード単位／ファイル名ラベル／ガイド／拡大・縮小、右：対象アートボード）
    // Lay out options in two columns (left: per-artboard / file-name label / guides / scale, right: target artboards)
    var optionsColumns = optionsPanel.add("group");
    setupGroup(optionsColumns, "row");
    optionsColumns.alignChildren = ["left", "top"]; // 2カラムを上端で揃える / Top-align the two columns
    optionsColumns.spacing = 20;                     // 左右カラムの間隔 / Gap between the two columns

    // 左カラム：アートボード単位／ファイル名ラベル／ガイド／拡大・縮小 / Left column
    var optionsLeftColumn = optionsColumns.add("group");
    setupGroup(optionsLeftColumn, "column");
    var byArtboardCheckbox = optionsLeftColumn.add("checkbox", undefined, getLocalizedText('checkbox.byArtboard'));
    byArtboardCheckbox.value = true;
    byArtboardCheckbox.helpTip = getLocalizedText('tooltip.byArtboard');

    var showLabelCheckbox = optionsLeftColumn.add("checkbox", undefined, getLocalizedText('checkbox.attachLabel'));
    showLabelCheckbox.value = openDocsRadio.value; // フォルダー指定では既定OFF / Default off in folder mode
    showLabelCheckbox.helpTip = getLocalizedText('tooltip.attachLabel');

    var includeGuidesCheckbox = optionsLeftColumn.add("checkbox", undefined, getLocalizedText('checkbox.includeGuides'));
    includeGuidesCheckbox.value = false;
    includeGuidesCheckbox.helpTip = getLocalizedText('tooltip.includeGuides');

    // スケール（チェックON時に％で拡大縮小）/ Scale (resize by percent when checked)
    var scaleRow = optionsLeftColumn.add("group");
    setupGroup(scaleRow, "row");
    var scaleCheckbox = scaleRow.add("checkbox", undefined, getLocalizedText('checkbox.scale'));
    scaleCheckbox.value = false;
    scaleCheckbox.helpTip = getLocalizedText('tooltip.scale');
    var scaleInput = scaleRow.add("edittext", undefined, "100");
    scaleInput.helpTip = getLocalizedText('tooltip.scale');
    scaleInput.characters = 4;
    scaleInput.enabled = scaleCheckbox.value;
    scaleRow.add("statictext", undefined, "%");
    scaleCheckbox.onClick = function () {
        scaleInput.enabled = this.value;
    };

    // 右カラム：対象アートボード（1のみ／すべて／指定）をパネルに / Right column: target artboards in a panel
    var targetArtboardPanel = optionsColumns.add("panel", undefined, getLocalizedText('field.artboardTarget'));
    setupPanel(targetArtboardPanel);
    var artboardOneRadio = targetArtboardPanel.add("radiobutton", undefined, getLocalizedText('radio.artboardOne'));
    artboardOneRadio.helpTip = getLocalizedText('tooltip.artboardTarget');
    var artboardAllRadio = targetArtboardPanel.add("radiobutton", undefined, getLocalizedText('radio.artboardAll'));
    artboardAllRadio.helpTip = getLocalizedText('tooltip.artboardTarget');
    var artboardSpecRow = targetArtboardPanel.add("group");
    setupGroup(artboardSpecRow, "row");
    var artboardSpecRadio = artboardSpecRow.add("radiobutton", undefined, getLocalizedText('radio.artboardSpecify'));
    artboardSpecRadio.helpTip = getLocalizedText('tooltip.artboardSpecify');
    var artboardSpecInput = artboardSpecRow.add("edittext", undefined, "");
    artboardSpecInput.characters = 7;
    artboardSpecInput.helpTip = getLocalizedText('tooltip.artboardSpecify');

    // 3つのラジオは別コンテナにまたがるため、排他選択を手動で制御する
    // The three radios span different containers, so enforce mutual exclusivity manually
    function selectArtboardTarget(which) {
        artboardOneRadio.value = (which === "one");
        artboardAllRadio.value = (which === "all");
        artboardSpecRadio.value = (which === "spec");
        artboardSpecInput.enabled = artboardSpecRadio.value;
    }
    artboardOneRadio.onClick = function () { selectArtboardTarget("one"); };
    artboardAllRadio.onClick = function () { selectArtboardTarget("all"); };
    artboardSpecRadio.onClick = function () { selectArtboardTarget("spec"); };
    selectArtboardTarget("one"); // 既定：アートボード1のみ / Default: Artboard 1 only

    // 対象アートボードは「アートボード単位」ONのときのみ有効（OFFでディム）
    // Target artboards apply only when "per artboard" is on (dimmed when off)
    function updateArtboardTargetState() {
        targetArtboardPanel.enabled = byArtboardCheckbox.value;
    }
    byArtboardCheckbox.onClick = updateArtboardTargetState;
    updateArtboardTargetState(); // 初期状態を反映 / Apply the initial state

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

    // ファイル名フィルターの正規表現を検証する。入力途中は無視してよいが、OK後は不正な式を黙って
    // 「全件対象」にせず中断する（誤って意図しないファイルまで取り込む事故を防ぐ）。
    // Validate the file-name filter regex. It's fine to ignore while typing, but after OK don't silently
    // fall back to "match all" on an invalid pattern — stop instead (avoids importing unintended files).
    if (filterInput.text !== "") {
        try {
            new RegExp(filterInput.text);
        } catch (eFilterPattern) {
            alert(getLocalizedText('alert.invalidFilter'));
            return;
        }
    }

    // 選択されたソースとオプションを確定 / Resolve the chosen source and options
    var importFromFolder = folderRadio.value;
    var useCurrentDoc = currentDocRadio.value; // 読み込み先：現在のドキュメント / Destination: current document
    var importByArtboard = byArtboardCheckbox.value;
    // アートボード対象：first（1のみ）/ all（すべて）/ specify（指定） / Target artboards
    var artboardTargetMode = artboardOneRadio.value ? "first" : (artboardAllRadio.value ? "all" : "specify");
    var artboardSpecText = artboardSpecInput.text;
    var includeGuides = includeGuidesCheckbox.value;
    var originalDocs = importFromFolder ? folderFiles : getFilteredOpenDocs();

    // 「現在のドキュメント」に読み込む場合は対象ドキュメントを確定する / Resolve the target document when importing into the current one
    var targetDoc = null;
    if (useCurrentDoc) {
        if (app.documents.length === 0) {
            alert(getLocalizedText('alert.noCurrentDoc'));
            return;
        }
        targetDoc = app.activeDocument;
        // 対象ドキュメント自身は取り込み対象から外す（誤って自分を開いて閉じるのを防ぐ）
        // Exclude the target document itself from the sources (avoid opening and then closing it)
        if (importFromFolder) {
            // フォルダー指定：同じファイルパスのファイルを除外（対象が保存済みファイルの場合のみ判定可能）
            // Folder mode: drop any file with the same path as the target (only when the target is a saved file)
            var targetPath = null;
            try {
                var targetFile = targetDoc.fullName;
                if (targetFile && targetFile.exists) targetPath = targetFile.fsName;
            } catch (ePath) { targetPath = null; }
            if (targetPath !== null) {
                var folderSourcesWithoutTarget = [];
                for (var fi = 0; fi < originalDocs.length; fi++) {
                    if (originalDocs[fi].fsName !== targetPath) folderSourcesWithoutTarget.push(originalDocs[fi]);
                }
                originalDocs = folderSourcesWithoutTarget;
            }
        } else {
            // 開いているファイル指定：対象ドキュメントの参照を除外 / Open-files mode: drop the target document reference
            var sourcesWithoutTarget = [];
            for (var od = 0; od < originalDocs.length; od++) {
                if (originalDocs[od] !== targetDoc) sourcesWithoutTarget.push(originalDocs[od]);
            }
            originalDocs = sourcesWithoutTarget;
        }
    }

    if (originalDocs.length < 1) {
        alert(getLocalizedText('alert.noValidFile'));
        return;
    }
    var originalDocsLength = originalDocs.length;

    // 「指定」モードでアートボード番号が一つも解釈できない場合は中断（無言終了を防ぐ）
    // Abort if "Specify" mode yields no parseable artboard numbers (avoids silently finishing)
    if (importByArtboard && artboardTargetMode === "specify" && parseArtboardNumbers(artboardSpecText).length === 0) {
        alert(getLocalizedText('alert.invalidArtboardSpec'));
        return;
    }

    // スケール（％）。チェックOFFなら100%（等倍）。ONで不正値（数値でない・0以下）は黙って等倍にせず中断する。
    // Scale percent; 100% when unchecked. When checked, abort on an invalid value (non-numeric or <= 0) instead of silently using 100%.
    var scalePercent = 100;
    if (scaleCheckbox.value) {
        var scaleValue = parseFloat(scaleInput.text);
        if (isNaN(scaleValue) || scaleValue <= 0) {
            alert(getLocalizedText('alert.invalidScale'));
            return;
        }
        scalePercent = scaleValue;
    }

    // 新規ドキュメント作成時のみ、寸法をプログレス表示の前に検証する（無効ならパレットを残さず終了）
    // For a new document only, validate the size before showing the progress palette (so it isn't left open on error)
    var docWidthValue, docHeightValue;
    if (!useCurrentDoc) {
        docWidthValue = parseFloat(widthInput.text);
        docHeightValue = parseFloat(heightInput.text);
        if (isNaN(docWidthValue) || isNaN(docHeightValue)) {
            alert(getLocalizedText('alert.invalidNumber'));
            return;
        }
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

    var newDoc;
    if (useCurrentDoc) {
        // 現在のドキュメントに読み込む（新規作成しない。カラーモード・解像度・サイズ設定は使わない）
        // Import into the current document (no new doc; color mode / resolution / size settings are not used)
        newDoc = targetDoc;
        app.activeDocument = newDoc;
    } else {
        // 入力値を現在の単位からポイントへ換算 / Convert the input values from the current unit to points
        var ptPerUnit = UNIT_TO_PT[currentUnit];
        var docWidthPt = docWidthValue * ptPerUnit;
        var docHeightPt = docHeightValue * ptPerUnit;
        var colorSpace = rgbRadio.value ? DocumentColorSpace.RGB : DocumentColorSpace.CMYK;

        // 通常ドキュメントの最大寸法（pt）。これを超える場合はラージカンバスとして作成。
        // Max dimension (pt) of a standard document; beyond this, create as a large canvas.
        var STANDARD_MAX_PT = 16383;

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

        // ラスタライズ効果設定の解像度を反映 / Apply the selected raster effects resolution
        var rasterPpi = parseInt(resolutionDropdown.selection.text, 10);
        if (!isNaN(rasterPpi)) {
            var rasterOptions = newDoc.rasterEffectSettings;
            rasterOptions.resolution = rasterPpi;
            newDoc.rasterEffectSettings = rasterOptions;
        }
    }

    // 配置の基準アートボード。新規は初期（仮）アートボード、現在のドキュメントはアクティブなアートボード。
    // Base artboard for placement: the placeholder for a new doc, the active artboard for the current doc.
    var baseArtboardIndex = useCurrentDoc ? newDoc.artboards.getActiveArtboardIndex() : 0;
    var artboardRect = newDoc.artboards[baseArtboardIndex].artboardRect; // [L, T, R, B]
    var canvasCenterX = (artboardRect[0] + artboardRect[2]) / 2;
    var canvasCenterY = (artboardRect[1] + artboardRect[3]) / 2;

    // 現在のドキュメントに読み込む場合は、取り込み前の既存アートボード全体を記録し、その下に重ならないよう配置する。
    // When importing into the current document, capture the existing artboards (before import) so the grid lands below them without overlapping.
    var avoidRect = useCurrentDoc ? getArtboardsUnionRect(newDoc) : null;

    // 配置の状態をまとめて保持。各セルを記録し、ループ後に正方形グリッドへ中央配置する。
    // Shared placement state. Each cell is recorded, then laid out into a centered square grid after the loop.
    var placementContext = {
        newDoc: newDoc,
        cells: [],
        canvasCenterX: canvasCenterX,
        canvasCenterY: canvasCenterY,
        avoidRect: avoidRect,
        byArtboard: importByArtboard,
        createArtboards: importByArtboard, // アートボードを作るのはアートボード単位ONのときだけ / Create artboards only in per-artboard mode
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
                    // アートボード単位：対象アートボードごとに、ロック・非表示も含めて取り込む
                    // Per artboard: import every object (incl. locked/hidden) on each target artboard
                    var targetIndices = resolveTargetArtboardIndices(srcDoc, artboardTargetMode, artboardSpecText);
                    // すべて：ドキュメント全体を解除。1のみ／指定：対象アートボードに重なるオブジェクトだけ解除し、他は触れない。
                    // 閉じない場合のみ後で復元する。
                    // All: unlock the whole doc. First/Specify: unlock only the objects overlapping the target
                    // artboards, leaving the rest untouched. Restore afterward only when keeping the file open.
                    if (artboardTargetMode === "all") {
                        if (!shouldCloseSource) lockState = captureLockHiddenState(srcDoc);
                        unlockAllLayersAndItems(srcDoc);
                    } else {
                        var targetRects = [];
                        for (var ti = 0; ti < targetIndices.length; ti++) {
                            targetRects.push(srcDoc.artboards[targetIndices[ti]].artboardRect);
                        }
                        var scopedState = unlockItemsOnArtboards(srcDoc, targetRects);
                        if (!shouldCloseSource) lockState = scopedState;
                    }
                    for (var ai = 0; ai < targetIndices.length; ai++) {
                        var ab = targetIndices[ai];
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
                    // 通常（アートボード単位OFF）：ファイル全体の表示中・選択可能オブジェクトを1つにまとめて取り込む。
                    // 「対象アートボード」はアートボード単位ONのときのみ有効なので、ここでは全選択する。
                    // Default (per-artboard off): import the whole file's visible/selectable objects as one.
                    // "Target artboards" applies only when per-artboard is on, so select everything here.
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

        // グループごとにアートボードを追加したので、初期の仮アートボードを削除（新規ドキュメント時のみ）。
        // 現在のドキュメントに読み込む場合は既存アートボードなので削除しない。
        // Each group added its own artboard, so remove the initial placeholder (new document only).
        // When importing into the current document, artboards[0] is the user's existing one — keep it.
        if (!useCurrentDoc && placementContext.placedCount > 0 && newDoc.artboards.length > placementContext.placedCount) {
            newDoc.artboards.remove(0);
        }
        app.activeDocument = newDoc;
    } finally {
        // ScriptUI パレットの close() は Illustrator で稀に "an Illustrator error occurred (MRAP)" を
        // 投げることがある（処理自体は完了済み）。spurious なエラーで全体を落とさないよう保護する。
        // また本体側で起きた本物の例外を close() のエラーで上書きしないためにも try/catch で囲む。
        // The palette's close() occasionally throws a spurious MRAP error in Illustrator even though the
        // work is done; guard it so it neither aborts the run nor masks a real error from the loop body.
        try { progressWin.hide(); } catch (hideErr) { }
        try { progressWin.close(); } catch (closeErr) { }
    }

    // すべてのアートボードがウィンドウに収まるように表示 / Fit all artboards in the window
    app.executeMenuCommand('fitall');

    // キャンセルされた場合はエラーではなく通常のメッセージで知らせる
    // If cancelled, inform the user with a normal message (not an error)
    if (userCancelled) {
        alert(getLocalizedText('alert.cancelled'));
    } else if (placementContext.placedCount === 0) {
        // 1件も配置できなかった場合（対象アートボード番号が存在しない・内容が空など）も無言で終わらせない
        // Don't finish silently when nothing was placed (e.g. target artboard numbers don't exist, or empty content)
        alert(getLocalizedText('alert.noArtboardImported'));
    }
})();