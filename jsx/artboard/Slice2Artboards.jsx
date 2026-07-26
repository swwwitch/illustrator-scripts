#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択した画像やオブジェクトを、指定した行数・列数のグリッドに分割し、各ピースを矩形マスクでクリッピングします。
- 分割したピースごとにアートボードを作成できます。
- アスペクト比のプリセット（A4、スクエア、16:9 など）を選ぶと、選択物の縦横比から行数・列数を自動で決めます。
- 印刷の面付け、パズル風レイアウト、複数アートボード化に使えます。
- 日本語／英語インターフェースに対応します。

### 主な機能

- 行数・列数の指定によるグリッド分割
- アスペクト比のプリセット選択（A4／スクエア／16:9／8:9／カスタム、英語環境では US Letter／US Legal／Tabloid も選択可）
- オフセットによるピースの拡大・縮小
- アートボードの自動生成、接頭辞・連番・ゼロ埋め・区切り記号の指定
- ファイル名をアートボード名に反映
- アートボード周囲のマージン指定
- 上下キーによる数値の増減（Shift：10単位）

### 処理の流れ

1. 選択オブジェクトをシンボル化し、マスク用の矩形を用意する
2. ダイアログで分割数・アスペクト比・アートボードの設定を行う
3. 実行でグリッド状の矩形マスクを生成し、各ピースをクリッピングする
4. 必要に応じてアートボードを作成・リネームし、元のアートボードを削除する
5. 元画像と一時的な矩形を削除する

### 注意

- オフセットとマージンの数値は定規単位のラベルを表示しますが、値はポイントとして扱われます。
- Illustrator には取り消しをグループ化する API がないため、この処理の取り消しは複数ステップに分かれます。

*/

/*

### Overview

- Splits the selected image or object into a grid of the given rows and columns, clipping each piece with a rectangular mask.
- Can create one artboard per piece.
- Picking an aspect-ratio preset (A4, Square, 16:9, and so on) derives the rows and columns from the aspect ratio of the selection.
- Useful for print imposition, puzzle-like layouts, and multi-artboard documents.
- Japanese and English user interface.

### Main Features

- Grid splitting by rows and columns
- Aspect-ratio presets (A4 / Square / 16:9 / 8:9 / Custom, plus US Letter / US Legal / Tabloid in the English UI)
- Offset to grow or shrink each piece
- Artboard creation with prefix, sequence, zero padding, and separator options
- File name reflected in the artboard name
- Margin around each artboard
- Arrow-key stepping (Shift: by 10)

### Workflow

1. Symbolize the selection and prepare the mask rectangle.
2. Configure divisions, aspect ratio, and artboard options in the dialog.
3. Run to generate the grid of rectangular masks and clip each piece.
4. Optionally create and rename artboards, then remove the original ones.
5. Remove the original image and the temporary rectangle.

### Notes

- The offset and margin fields show the ruler unit, but the values are treated as points.
- Illustrator has no undo-grouping API, so undoing this run takes multiple steps.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "Slice2Artboards";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-10";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-27";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/Slice2Artboards.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/Slice2Artboards.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User configuration
    // =========================================
    var CONFIG = {
        defaultColumnCount: 5,     /* 列数の初期値 / default columns */
        defaultRowCount: 5,        /* 行数の初期値 / default rows */
        defaultOffset: -20,        /* オフセットの初期値 / default offset */
        defaultMargin: 0,          /* マージンの初期値 / default margin */
        defaultStartNumber: 1,     /* 連番の初期値 / default sequence start */
        fallbackFileName: "Artboard" /* ファイル名が取れないときの代替 / fallback when the file name is unavailable */
    };

    /* アスペクト比のプリセット（ratio が null のカラムはユーザー入力優先）/ Aspect-ratio presets; a null ratio means the user's input wins */
    var SHAPE_PRESETS = [
        { key: "a4", ratio: 210 / 297 },
        { key: "square", ratio: 1.0 },
        { key: "letter", ratio: 8.5 / 11, englishOnly: true },
        { key: "legal", ratio: 8.5 / 14, englishOnly: true },
        { key: "tabloid", ratio: 11 / 17, englishOnly: true },
        { key: "ratio169", ratio: 16 / 9 },
        { key: "ratio89", ratio: 8 / 9 },
        { key: "custom", ratio: null }
    ];

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 実行環境のロケールから表示言語を決める / Pick the UI language from the locale */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "分割してアートボード化", en: "Slice and Create Artboards" }
        },
        panel: {
            shape: { ja: "アスペクト比", en: "Aspect Ratio" },
            artboard: { ja: "アートボード名", en: "Artboard Name" }
        },
        shape: {
            a4: { ja: "A4 (210 x 297)", en: "A4 (210 x 297)" },
            square: { ja: "スクエア", en: "Square" },
            letter: { ja: "US Letter", en: "US Letter" },
            legal: { ja: "US Legal (8.5 x 14)", en: "US Legal (8.5 x 14)" },
            tabloid: { ja: "Tabloid (11 x 17)", en: "Tabloid (11 x 17)" },
            ratio169: { ja: "16:9", en: "16:9" },
            ratio89: { ja: "8:9", en: "8:9" },
            custom: { ja: "カスタム", en: "Custom" }
        },
        field: {
            columns: { ja: "列数", en: "Columns" },
            rows: { ja: "行数", en: "Rows" },
            offset: { ja: "オフセット", en: "Offset" },
            prefix: { ja: "接頭辞", en: "Prefix" },
            startNumber: { ja: "開始番号", en: "Starting Number" },
            separator: { ja: "記号", en: "Separator" },
            margin: { ja: "マージン", en: "Margin" }
        },
        checkbox: {
            convertArtboard: { ja: "アートボードに変換", en: "Convert to Artboards" },
            zeroPad: { ja: "ゼロ埋め", en: "Zero Padding" },
            useFileName: { ja: "ファイル名を参照", en: "Use File Name" }
        },
        separator: {
            dash: { ja: "-", en: "-" },
            underscore: { ja: "_", en: "_" },
            without: { ja: "なし", en: "None" }
        },
        button: {
            run: { ja: "実行", en: "Run" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        message: {
            noDocument: {
                ja: "ドキュメントを開いてください。",
                en: "Please open a document."
            },
            noSelection: {
                ja: "オブジェクトを選択してから実行してください。",
                en: "Please select an object before running this script."
            },
            runError: {
                ja: "スクリプトの実行中にエラーが発生しました：\n",
                en: "An error occurred while running the script:\n"
            },
            symbolizeGroupError: {
                ja: "複数オブジェクトのシンボル化に失敗しました：\n",
                en: "Failed to symbolize the selected objects:\n"
            },
            symbolizeItemError: {
                ja: "オブジェクトのシンボル化に失敗しました：\n",
                en: "Failed to symbolize the object:\n"
            }
        }
    };

    /* ラベルをドット区切りのキーで引く / Look a label up by a dot-separated key */
    function L(key) {
        var parts = key.split(".");
        var entry = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (!entry) return key;
            entry = entry[parts[i]];
        }
        return (entry && entry[currentLanguage]) ? entry[currentLanguage] : key;
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(key) {
        return L(key) + (currentLanguage === "ja" ? "：" : ":");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* rulerType の並びに対応した単位ラベル / Unit labels matching the rulerType order */
    var UNIT_LABELS = ["in", "mm", "pt", "pica", "cm", "Q/H", "px", "ft/in", "m", "yd", "ft"];

    /* 現在の定規単位のラベルを返す / Return the label of the current ruler unit */
    function getCurrentUnitLabel() {
        var unitCode = app.preferences.getIntegerPreference("rulerType");
        return UNIT_LABELS[unitCode] || "pt";
    }

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

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

    /* 行グループの共通設定（ボタン列など） / Apply a horizontal row group */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignment = alignment || "left";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ）/ Trim a button's height by the given px (call after layout) */
    function trimButtonHeight(button, px) {
        try {
            button.size = [button.size.width, button.size.height - px];
        } catch (e) { }
    }

    /* 上下キーで数値を増減する（Shift：10単位）/ Step a value with the arrow keys (Shift: by 10) */
    function enableArrowKeyStep(inputField, allowsNegative) {
        inputField.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = Number(inputField.text);
            if (isNaN(value)) return;

            var direction = (event.keyName === "Up") ? 1 : -1;
            var step = 1;

            if (ScriptUI.environment.keyboardState.shiftKey) {
                /* 10 の倍数へ丸めてから増減 / Snap to a multiple of 10 before stepping */
                value = Math.floor(value / 10) * 10;
                step = 10;
            }

            value += direction * step;
            if (!allowsNegative && value < 0) value = 0;

            inputField.text = String(value);
            event.preventDefault();
        });
    }

    // =========================================
    // 幾何の計算 / Geometry helpers
    // =========================================

    /* 矩形の幅と高さを返す（高さは常に正）/ Return the width and height of a rect, height always positive */
    function getSize(bounds) {
        var height = bounds[1] - bounds[3];
        return {
            width: bounds[2] - bounds[0],
            height: (height < 0) ? -height : height
        };
    }

    /* 矩形をオフセット分だけ拡大・縮小する / Grow or shrink a rectangle by the offset */
    function offsetRectangle(pathItem, offsetX, offsetY) {
        var size = getSize(pathItem.geometricBounds);
        pathItem.left = pathItem.left - offsetX;
        pathItem.top = pathItem.top + offsetY;
        pathItem.width = size.width + offsetX * 2;
        pathItem.height = size.height + offsetY * 2;
        return pathItem;
    }

    /* 選択物の縦横比とプリセットから、行数・列数を求める / Derive rows and columns from the selection's aspect ratio */
    function computeDivision(size, targetRatio) {
        var aspect = size.width / size.height;
        if (aspect > targetRatio) {
            return { rowCount: 1, columnCount: Math.max(1, Math.round(aspect / targetRatio)) };
        }
        return { columnCount: 1, rowCount: Math.max(1, Math.round(targetRatio / aspect)) };
    }

    /* 1ピースの寸法を求める（カスタムは等分割、それ以外は比率に合わせる）/ Size of one piece: equal division for Custom, ratio-driven otherwise */
    function computePieceSize(size, columnCount, rowCount, targetRatio) {
        var cellWidth = size.width / columnCount;
        var cellHeight = size.height / rowCount;

        if (targetRatio === null) return { width: cellWidth, height: cellHeight };

        if (cellWidth / cellHeight > targetRatio) {
            return { width: cellHeight * targetRatio, height: cellHeight };
        }
        return { width: cellWidth, height: cellWidth / targetRatio };
    }

    // =========================================
    // シンボル化 / Symbolize
    // =========================================

    /* オブジェクトをシンボルに置き換え、元の位置を保つ / Replace an item with a symbol instance, keeping its position */
    function replaceWithSymbol(doc, item) {
        var left = item.left;
        var top = item.top;
        var parent = item.parent;

        var symbolItem = parent.symbolItems.add(doc.symbols.add(item));
        symbolItem.left = left;
        symbolItem.top = top;
        item.remove();
        return symbolItem;
    }

    /* 複数選択をグループ化してから1つのシンボルにまとめる（重ね順を保持）/ Group a multi-selection into a single symbol, preserving the stacking order */
    function symbolizeSelection(doc, selection) {
        var group = doc.groupItems.add();
        var items = [];
        for (var i = 0; i < selection.length; i++) items.push(selection[i]);

        /* 後ろのものから移動すると重ね順が保たれる / Moving from the back preserves the order */
        for (var j = items.length - 1; j >= 0; j--) {
            items[j].move(group, ElementPlacement.PLACEATBEGINNING);
        }
        return replaceWithSymbol(doc, group);
    }

    /* オブジェクトの外形と同じ矩形を作る（マスク用）/ Build a rectangle matching the item's bounds, used as the mask */
    function createBoundsRectangle(doc, item) {
        var bounds = item.geometricBounds;
        var size = getSize(bounds);
        return doc.pathItems.rectangle(bounds[1], bounds[0], size.width, size.height);
    }

    /*
     * 選択をシンボル化し、分割の基準となる矩形を用意する。
     * リンク画像は複製してから元を削除し、埋め込み画像やベクターはシンボルに変換する。
     * Symbolize the selection and prepare the rectangle the grid is based on.
     * Linked images are duplicated and the original removed; embedded images and vectors become symbols.
     */
    function prepareSource(doc, selection) {
        var result = { symbolItem: null, maskRect: null, origImageObj: null, isTempRect: false };
        var item;

        try {
            item = (selection.length > 1) ? symbolizeSelection(doc, selection) : selection[0];
        } catch (e) {
            alert(L("message.symbolizeGroupError") + e);
            return result;
        }

        /* リンク画像は複製して元を削除する（マスクは複製側に掛ける）/ Duplicate a linked image and drop the original */
        if (item.typename === "PlacedItem" && !item.embedded) {
            var placedCopy = item.duplicate();
            item.remove();
            result.symbolItem = placedCopy;
            result.maskRect = createBoundsRectangle(doc, placedCopy);
            result.origImageObj = placedCopy;
            result.isTempRect = true;
            return result;
        }

        /* 埋め込み画像・ベクターはシンボルに変換してから扱う / Convert embedded images and vectors to symbols first */
        if (item.typename === "RasterItem" ||
            item.typename === "PathItem" ||
            item.typename === "GroupItem" ||
            item.typename === "CompoundPathItem") {
            try {
                item = replaceWithSymbol(doc, item);
            } catch (e) {
                alert(L("message.symbolizeItemError") + e);
                return result;
            }
        }

        if (item.typename === "PlacedItem" || item.typename === "RasterItem" || item.typename === "SymbolItem") {
            result.maskRect = createBoundsRectangle(doc, item);
            result.origImageObj = item;
            result.isTempRect = true;
        }
        result.symbolItem = item;
        return result;
    }

    // =========================================
    // グリッドの生成 / Grid generation
    // =========================================

    /*
     * 行数・列数のぶんだけ矩形マスクを作り、画像やシンボルならクリッピングする。
     * Create one rectangular mask per cell, clipping the image or symbol when there is one.
     */
    function createGridPieces(doc, origin, pieceSize, columnCount, rowCount, offsetValue, sourceItem) {
        var pieces = [];
        var clipsSource = sourceItem &&
            (sourceItem.typename === "PlacedItem" || sourceItem.typename === "SymbolItem");

        for (var row = 0; row < rowCount; row++) {
            for (var column = 0; column < columnCount; column++) {
                var mask = doc.pathItems.rectangle(
                    origin.top - row * pieceSize.height,
                    origin.left + column * pieceSize.width,
                    pieceSize.width,
                    pieceSize.height
                );

                if (offsetValue !== null) mask = offsetRectangle(mask, offsetValue, offsetValue);
                mask.closed = true;
                mask.filled = false;
                mask.stroked = false;

                if (!clipsSource) {
                    pieces.push(mask);
                    continue;
                }

                var group = doc.groupItems.add();
                sourceItem.duplicate().moveToBeginning(group);
                mask.moveToBeginning(group);
                mask.clipping = true;
                group.clipped = true;
                pieces.push(group);
            }
        }

        /* 生成順と重ね順を合わせる / Match the stacking order to the generation order */
        for (var i = 0; i < pieces.length; i++) {
            pieces[i].zOrder(ZOrderMethod.SENDTOBACK);
        }
        return pieces;
    }

    /* ピースからマスクの矩形を取り出す / Pull the mask rectangle out of a piece */
    function getMaskRect(piece) {
        if (piece.typename === "PathItem") return piece;
        if (piece.typename !== "GroupItem" || !piece.clipped) return null;

        for (var i = 0; i < piece.pageItems.length; i++) {
            if (piece.pageItems[i].clipping) return piece.pageItems[i];
        }
        return null;
    }

    /* ピースを見た目の順（上から下、左から右）に並べ替える / Sort pieces in visual order: top to bottom, left to right */
    function sortPiecesByPosition(pieces) {
        var sorted = [];
        for (var i = 0; i < pieces.length; i++) {
            var maskRect = getMaskRect(pieces[i]);
            if (maskRect) sorted.push({ piece: pieces[i], bounds: maskRect.geometricBounds });
        }

        /* Illustrator の座標系は上ほど値が大きい / In Illustrator's coordinates, higher means further up */
        sorted.sort(function (a, b) {
            if (a.bounds[1] !== b.bounds[1]) return b.bounds[1] - a.bounds[1];
            return a.bounds[0] - b.bounds[0];
        });
        return sorted;
    }

    // =========================================
    // アートボードの生成 / Artboard generation
    // =========================================

    /* アートボード名を組み立てる / Build the artboard name */
    function buildArtboardName(options, sequence) {
        var numberText = String(sequence);
        while (options.zeroPadding && numberText.length < options.digitCount) {
            numberText = "0" + numberText;
        }

        var parts = [];
        if (options.useFileName) parts.push(options.fileName);
        if (options.prefix !== "") parts.push(options.prefix);
        parts.push(numberText);
        return parts.join(options.separator);
    }

    /* 現在のファイル名（拡張子なし）を返す / Return the current file name without its extension */
    function getFileNameWithoutExtension(doc) {
        return doc.name.replace(/\.[^\.]+$/, "") || CONFIG.fallbackFileName;
    }

    /*
     * 並べ替えたピースごとにアートボードを作り、既存のアートボードを削除する。
     * Create one artboard per sorted piece, then remove the pre-existing artboards.
     */
    function createArtboardsFromPieces(doc, sortedPieces, options) {
        var existingCount = doc.artboards.length;

        for (var i = 0; i < sortedPieces.length; i++) {
            var bounds = sortedPieces[i].bounds;
            doc.artboards.add([
                bounds[0] - options.margin,
                bounds[1] + options.margin,
                bounds[2] + options.margin,
                bounds[3] - options.margin
            ]);

            var name = buildArtboardName(options, options.startNumber + i);
            doc.artboards[doc.artboards.length - 1].name = name;
            sortedPieces[i].piece.name = name;
        }

        /* 分割前からあったアートボードを後ろから削除する / Remove the pre-existing artboards, back to front */
        for (var j = existingCount - 1; j >= 0; j--) {
            doc.artboards.remove(j);
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /*
     * 分割設定のダイアログを組み立てる。
     * 戻り値の read() が、入力内容をまとめたオブジェクトを返す。
     * Build the settings dialog. The returned read() collects the entered values.
     */
    function buildDialog(sourceSize) {
        var dialog = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        var mainGroup = dialog.add("group");
        setupRow(mainGroup, "fill", COLUMN_SPACING);
        mainGroup.alignChildren = ["fill", "top"];

        var leftColumn = mainGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];

        var rightColumn = mainGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "top"];

        /* アスペクト比パネル（プリセット表からラジオボタンを生成）/ Aspect-ratio panel, built from the preset table */
        var shapePanel = leftColumn.add("panel", undefined, L("panel.shape"));
        setupPanel(shapePanel, 6);
        shapePanel.alignChildren = ["left", "top"];

        var shapeRadios = [];
        for (var i = 0; i < SHAPE_PRESETS.length; i++) {
            var preset = SHAPE_PRESETS[i];
            if (preset.englishOnly && currentLanguage === "ja") continue;

            var radio = shapePanel.add("radiobutton", undefined, L("shape." + preset.key));
            radio.alignment = "left";
            shapeRadios.push({ radio: radio, ratio: preset.ratio });
        }
        shapeRadios[0].radio.value = true;

        /* 行数・列数 / Rows and columns */
        var divisionRow = leftColumn.add("group");
        setupRow(divisionRow, "left", COLUMN_SPACING);

        divisionRow.add("statictext", undefined, labelText("field.columns"));
        var columnInput = divisionRow.add("edittext", undefined, String(CONFIG.defaultColumnCount));
        columnInput.characters = 3;

        divisionRow.add("statictext", undefined, labelText("field.rows"));
        var rowInput = divisionRow.add("edittext", undefined, String(CONFIG.defaultRowCount));
        rowInput.characters = 3;

        /* オフセット / Offset */
        var offsetRow = leftColumn.add("group");
        setupRow(offsetRow, "left", 6);

        var offsetCheckbox = offsetRow.add("checkbox", undefined, L("field.offset"));
        offsetCheckbox.value = true;
        var offsetInput = offsetRow.add("edittext", undefined, String(CONFIG.defaultOffset));
        offsetInput.characters = 4;
        var offsetUnitLabel = offsetRow.add("statictext", undefined, getCurrentUnitLabel());

        /* アートボード変換 / Convert to artboards */
        var artboardCheckbox = rightColumn.add("checkbox", undefined, L("checkbox.convertArtboard"));
        artboardCheckbox.value = true;
        artboardCheckbox.alignment = "left";

        var artboardPanel = rightColumn.add("panel", undefined, L("panel.artboard"));
        setupPanel(artboardPanel, 6);
        artboardPanel.alignChildren = ["left", "top"];

        var useFileNameCheckbox = artboardPanel.add("checkbox", undefined, L("checkbox.useFileName"));

        var prefixRow = artboardPanel.add("group");
        setupRow(prefixRow, "left", 6);
        prefixRow.add("statictext", undefined, labelText("field.prefix"));
        var prefixInput = prefixRow.add("edittext", undefined, "");
        prefixInput.characters = 14;

        var separatorRow = artboardPanel.add("group");
        setupRow(separatorRow, "left", 6);
        separatorRow.add("statictext", undefined, labelText("field.separator"));
        var dashRadio = separatorRow.add("radiobutton", undefined, L("separator.dash"));
        var underscoreRadio = separatorRow.add("radiobutton", undefined, L("separator.underscore"));
        var withoutRadio = separatorRow.add("radiobutton", undefined, L("separator.without"));
        dashRadio.value = true;

        var numberRow = artboardPanel.add("group");
        setupRow(numberRow, "left", 6);
        numberRow.add("statictext", undefined, labelText("field.startNumber"));
        var startNumberInput = numberRow.add("edittext", undefined, String(CONFIG.defaultStartNumber));
        startNumberInput.characters = 3;
        var zeroPadCheckbox = numberRow.add("checkbox", undefined, L("checkbox.zeroPad"));
        zeroPadCheckbox.value = true;

        var marginRow = rightColumn.add("group");
        setupRow(marginRow, "left", 6);
        marginRow.add("statictext", undefined, labelText("field.margin"));
        var marginInput = marginRow.add("edittext", undefined, String(CONFIG.defaultMargin));
        marginInput.characters = 5;
        marginRow.add("statictext", undefined, getCurrentUnitLabel());

        /* ボタン列 / Button row */
        var buttonRow = dialog.add("group");
        setupRow(buttonRow, "right", 10);
        var cancelButton = buttonRow.add("button", undefined, L("button.cancel"), { name: "cancel" });
        var runButton = buttonRow.add("button", undefined, L("button.run"), { name: "ok" });
        runButton.active = true;

        // -----------------------------------------
        // UI 状態の同期 / UI state
        // -----------------------------------------

        /* 選択中のアスペクト比を返す（カスタムは null）/ Return the selected aspect ratio; null for Custom */
        function getSelectedRatio() {
            for (var i = 0; i < shapeRadios.length; i++) {
                if (shapeRadios[i].radio.value) return shapeRadios[i].ratio;
            }
            return null;
        }

        /* プリセットに合わせて行数・列数を計算し直す / Recalculate rows and columns for the chosen preset */
        function syncDivision() {
            var ratio = getSelectedRatio();
            if (!sourceSize || ratio === null) return;

            var division = computeDivision(sourceSize, ratio);
            columnInput.text = String(division.columnCount);
            rowInput.text = String(division.rowCount);
        }

        /* オフセット欄の有効・無効を切り替える / Enable or disable the offset field */
        function syncOffsetEnabled() {
            offsetInput.enabled = offsetCheckbox.value;
            offsetUnitLabel.enabled = offsetCheckbox.value;
        }

        /* アートボード関連の入力欄をまとめて切り替える / Toggle the artboard-related fields together */
        function syncArtboardEnabled() {
            var enabled = artboardCheckbox.value;
            useFileNameCheckbox.enabled = enabled;
            prefixInput.enabled = enabled;
            startNumberInput.enabled = enabled;
            zeroPadCheckbox.enabled = enabled;
            marginInput.enabled = enabled;
            dashRadio.enabled = enabled;
            underscoreRadio.enabled = enabled;
            withoutRadio.enabled = enabled;
        }

        // -----------------------------------------
        // イベント / Event wiring
        // -----------------------------------------
        for (var k = 0; k < shapeRadios.length; k++) {
            shapeRadios[k].radio.onClick = syncDivision;
        }

        /* 行数・列数を直接触ったらカスタム扱いにする / Switch to Custom once the counts are edited by hand */
        function selectCustom() {
            shapeRadios[shapeRadios.length - 1].radio.value = true;
        }
        columnInput.onChanging = selectCustom;
        rowInput.onChanging = selectCustom;

        offsetCheckbox.onClick = syncOffsetEnabled;
        artboardCheckbox.onClick = syncArtboardEnabled;

        enableArrowKeyStep(columnInput, false);
        enableArrowKeyStep(rowInput, false);
        enableArrowKeyStep(offsetInput, true);
        enableArrowKeyStep(startNumberInput, false);
        enableArrowKeyStep(marginInput, true);

        syncDivision();
        syncOffsetEnabled();
        syncArtboardEnabled();

        dialog.layout.layout(true);
        trimButtonHeight(runButton, 0);

        /* 入力内容をまとめて取り出す / Collect the entered values */
        function read() {
            var startNumberText = startNumberInput.text;
            var startNumber = parseInt(startNumberText, 10);
            if (isNaN(startNumber)) startNumber = CONFIG.defaultStartNumber;

            var margin = parseFloat(marginInput.text);
            if (isNaN(margin)) margin = 0;

            var offsetValue = null;
            if (offsetCheckbox.value) {
                offsetValue = parseFloat(offsetInput.text);
                if (isNaN(offsetValue)) offsetValue = null;
            }

            return {
                columnCount: Math.round(Number(columnInput.text)),
                rowCount: Math.round(Number(rowInput.text)),
                targetRatio: getSelectedRatio(),
                offsetValue: offsetValue,
                createsArtboards: artboardCheckbox.value,
                prefix: prefixInput.text,
                separator: dashRadio.value ? "-" : (underscoreRadio.value ? "_" : ""),
                startNumber: startNumber,
                startNumberText: startNumberText,
                zeroPadding: zeroPadCheckbox.value,
                useFileName: useFileNameCheckbox.value,
                margin: margin
            };
        }

        return { window: dialog, read: read };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* 行数・列数が 0 のとき、選択物の縦横比から補う / Fill in a zero row or column count from the aspect ratio */
    function resolveCounts(settings, size) {
        var columnCount = settings.columnCount;
        var rowCount = settings.rowCount;

        if (columnCount === 0 && rowCount > 0) {
            columnCount = Math.max(1, Math.round((size.width / size.height) * rowCount));
        }
        if (rowCount === 0 && columnCount > 0) {
            rowCount = Math.max(1, Math.round((size.height / size.width) * columnCount));
        }
        return { columnCount: columnCount, rowCount: rowCount };
    }

    /* 分割からアートボード生成までを実行する / Run everything from the split to the artboards */
    function runDistribution(doc, settings, source) {
        var baseItem = source.maskRect ? source.maskRect : source.symbolItem;
        var bounds = baseItem.geometricBounds;
        var size = getSize(bounds);

        var counts = resolveCounts(settings, size);
        if (counts.columnCount < 1 || counts.rowCount < 1) return;

        baseItem.selected = false;

        var pieceSize = computePieceSize(size, counts.columnCount, counts.rowCount, settings.targetRatio);
        var pieces = createGridPieces(
            doc,
            { left: bounds[0], top: bounds[1] },
            pieceSize,
            counts.columnCount,
            counts.rowCount,
            settings.offsetValue,
            source.origImageObj
        );

        if (settings.createsArtboards) {
            var sortedPieces = sortPiecesByPosition(pieces);
            var maxNumber = settings.startNumber + sortedPieces.length - 1;

            createArtboardsFromPieces(doc, sortedPieces, {
                prefix: settings.prefix,
                separator: settings.separator,
                startNumber: settings.startNumber,
                zeroPadding: settings.zeroPadding,
                digitCount: settings.zeroPadding ? String(maxNumber).length : settings.startNumberText.length,
                useFileName: settings.useFileName,
                fileName: getFileNameWithoutExtension(doc),
                margin: settings.margin
            });
        }

        /* 分割の元にした画像と一時矩形を片付ける / Clean up the source image and the temporary rectangle */
        if (source.origImageObj) source.origImageObj.remove();
        if (source.isTempRect && source.maskRect) source.maskRect.remove();
    }

    /* 選択の確認、ダイアログ表示、分割の実行までを通して行う / Check the selection, show the dialog, then run the split */
    function main() {
        var doc = app.activeDocument;
        var selection = doc.selection;

        var sourceSize = (selection.length === 1) ? getSize(selection[0].geometricBounds) : null;
        var dialog = buildDialog(sourceSize);
        if (dialog.window.show() !== 1) return;

        var settings = dialog.read();
        if (settings.columnCount < 1 && settings.rowCount < 1) return;

        try {
            var source = prepareSource(doc, selection);
            if (!source.symbolItem) return;
            runDistribution(doc, settings, source);
        } catch (e) {
            alert(L("message.runError") + e);
        }
    }

    if (app.documents.length === 0) {
        alert(L("message.noDocument"));
    } else if (!app.activeDocument.selection || app.activeDocument.selection.length === 0) {
        alert(L("message.noSelection"));
    } else {
        main();
    }

})();
