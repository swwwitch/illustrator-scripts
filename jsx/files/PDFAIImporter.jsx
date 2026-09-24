#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

PDF/AI ファイルを指定したページ範囲で読み込み、現在のドキュメント上に配置します。
各ページを個別のアートボードとして並べるか、アートボードを追加せずオブジェクトとして配置するかを選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PDFAIImporter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n42595650216f

### Overview

Imports a PDF/AI file over a given page range and places the pages in the current document.
The pages can be laid out as one artboard each, or placed as objects without adding artboards.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PDFAIImporter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PDFAIImporter";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-13";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PDFAIImporter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PDFAIImporter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n42595650216f"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    // アートボードごと配置時のデフォルト間隔（pt）/ Default gap for per-artboard placement (pt)
    var DEFAULT_ARTBOARD_GAP = 100;

    // 自動列のときの最大行幅（pt）。これを超えると次行へ折り返す / Max row width for Auto columns (pt); wraps beyond this
    var MAX_AUTO_ROW_WIDTH = (220 / 2) * 72; // 7920 pt

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在の UI 言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* 日英ラベル定義 */

    var LABELS = {
        dialog: {
            title: { ja: "PDF{slash}AI配置", en: "PDF{slash}AI Placement" },
            pickFile: { ja: "PDF{slash}AIを選択してください", en: "Select a PDF{slash}AI" }
        },
        panel: {
            source: { ja: "読み込みファイル", en: "Source File" },
            pages: { ja: "対象ページ", en: "Pages" },
            mode: { ja: "配置方法", en: "Placement Method" },
            placement: { ja: "レイアウト", en: "Layout" },
            option: { ja: "オプション", en: "Options" },
            kei: { ja: "枠線", en: "Stroke" }
        },
        radio: {
            rangeAll: { ja: "全ページ", en: "All Pages" },
            rangeFirst: { ja: "先頭ページのみ", en: "First Page Only" },
            rangeCustom: { ja: "指定ページ", en: "Custom Pages" },
            perArtboard: { ja: "アートボードごと", en: "Per Artboard" },
            ignoreArtboard: { ja: "アートボードを無視", en: "Place as Objects" },
            keiNone: { ja: "なし", en: "None" },
            keiClipGroup: { ja: "枠線を追加", en: "Add stroke" }
        },
        checkbox: {
            roundCorner: { ja: "角丸", en: "Round corners" }
        },
        label: {
            notSelected: { ja: "未指定", en: "Not selected" },
            scale: { ja: "倍率", en: "Scale" },
            scaleUnit: { ja: "%", en: "%" },
            gap: { ja: "間隔", en: "Gap" },
            gapUnit: { ja: "pt", en: "pt" },
            columns: { ja: "列数", en: "Columns" },
            columnsAuto: { ja: "自動", en: "Auto" },
            rows: { ja: "行数", en: "Rows" },
            errorDetails: { ja: "詳細", en: "Details" }
        },
        button: {
            selectFile: { ja: "ファイル指定", en: "Select File" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            source: { ja: "PDF または AI ファイルを選択（複数ページ対応）", en: "Choose a PDF or AI file (multi-page supported)" },
            customRange: { ja: "例: 1-10, 1,3,5", en: "e.g. 1-10, 1,3,5" },
            totalPages: { ja: "配置されるページ数 {slash} 総ページ数", en: "Pages to place {slash} total pages" },
            columns: { ja: "1 行あたりの列数。自動はカンバス右端で折り返し", en: "Columns per row; Auto wraps at the canvas edge" },
            estimate: { ja: "配置に必要な行数（列数が自動のときは不定）", en: "Rows needed for the layout (unknown when Columns is Auto)" },
            gapPerArtboard: { ja: "アートボードの間隔", en: "Artboard gap" },
            gapIgnoreArtboard: { ja: "配置するオブジェクトの間隔", en: "Placed object gap" },
            scale: { ja: "配置倍率（［アートボードを無視］のときのみ有効）", en: "Placement scale (only when ignoring artboards)" },
            perArtboard: { ja: "各ページをアートボードとして並べる（倍率 100%）", en: "Lay out each page as an artboard (100%)" },
            ignoreArtboard: { ja: "アートボードを追加せずオブジェクトとして配置", en: "Place as objects without adding artboards" },
            roundCorner: { ja: "外接矩形の角を丸める", en: "Round the corners of the bounding rectangle" }
        },
        alert: {
            needDoc: {
                ja: "ドキュメントを開いてから実行してください。",
                en: "Please open a document before running."
            },
            needFile: {
                ja: "先に［ファイル指定］で読み込みファイルを選択してください。",
                en: "Please select a source file first."
            },
            placeError: {
                ja: "配置中にエラーが発生しました。",
                en: "An error occurred while placing the pages."
            },
            linkUnknown: {
                ja: "画像のリンク先が不明でした。",
                en: "Image link not found."
            },
            pageCountFail: {
                ja: "リンクされたPDF{slash}AIファイルのページ数を取得できませんでした。",
                en: "Could not determine the page count of the linked PDF{slash}AI file."
            },
            pickPdfAi: {
                ja: "PDFまたはAIファイルを選択してください。",
                en: "Please select a PDF or AI file."
            },
            someSkipped: {
                ja: "一部のページを配置できなかったため、スキップしました。",
                en: "Some pages could not be placed and were skipped."
            }
        }
    };

    /**
     * ラベル（ja/en のリーフ）を現在の言語に解決し、{slash} を / に置換する
     * @param {object} entry - ja / en を持つラベル定義
     * @returns {string} 表示する文言
     */
    function getLabel(entry) {
        if (!entry) return "";
        var text = entry[currentLanguage] || entry.en || entry.ja || "";
        return String(text).replace(/\{slash\}/g, "/");
    }

    /**
     * 項目名にコロンを付与する（日本語は全角、英語は半角）
     * @param {object} entry - ja / en を持つラベル定義
     * @returns {string} コロン付きの項目名
     */
    function labelText(entry) {
        return getLabel(entry) + (currentLanguage === "ja" ? "：" : ":");
    }

    // =========================================
    // 単位 / Unit
    // =========================================

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

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;

    /**
     * パネルへ共通のレイアウト設定を適用する
     * @param {object} panel - 対象のパネル
     * @param {number} spacing - 子要素の間隔。省略時は既定値
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
     * グループへ共通のレイアウト設定を適用する（row/column で整列を切り替え）
     * @param {object} group - 対象のグループ
     * @param {string} orientation - "row" または "column"
     * @param {number} spacing - 子要素の間隔。省略時は既定値
     * @returns {void}
     */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 例外オブジェクトから表示用の詳細文字列を取り出す
     * @param {object} e - 例外オブジェクトまたは文字列
     * @returns {string} 表示用の詳細文字列
     */
    function SC_getErrorDetailText(e) {
        if (e === undefined || e === null) return "";
        try {
            if (typeof e === "string") return e;
            if (e && e.message) return String(e.message);
            return String(e);
        } catch (err) {
            return "";
        }
    }

    /**
     * ラベルエントリと任意のエラーをアラート表示する
     * @param {object} entry - ja / en を持つラベル定義
     * @param {object} e - 併記する例外オブジェクト。省略可
     * @returns {void}
     */
    function SC_alert(entry, e) {
        try {
            var msg = getLabel(entry);
            var detail = SC_getErrorDetailText(e);
            if (detail) msg += "\n\n" + labelText(LABELS.label.errorDetails) + "\n" + detail;
            alert(msg);
        } catch (err) { }
    }

    // ========================
    // 配置処理ヘルパー
    // - 配置前のページ番号指定
    // - 配置サイズの計測
    // - レイアウト計算
    // - アートボードの作成 / 更新
    // - 単ページの配置
    // - 指定範囲が見える表示倍率への調整
    // ========================

    /**
     * PDF 取り込みページ番号を環境設定にセットする
     * @param {number} pageNum - 取り込むページ番号
     * @returns {void}
     */
    function placementSetImportPageNumber(pageNum) {
        var n = parseInt(pageNum, 10);
        if (isNaN(n) || n < 1) n = 1;
        try {
            app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", n);
        } catch (e) { }
    }

    /**
     * 取り込みページ番号を 1 に戻す
     * @returns {void}
     */
    function placementResetImportPageNumber() {
        try {
            app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", 1);
        } catch (e) { }
    }

    /**
     * 指定ページを一時配置してサイズを計測する
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {File} fileObj - 読み込み元ファイル
     * @param {number} pageNum - 計測するページ番号
     * @param {number} cropMode - トリミング指定（CropTo の値）
     * @returns {object} width と height を持つオブジェクト
     */
    function placementMeasurePlacedPageSize(targetDoc, fileObj, pageNum, cropMode) {
        if (isPdfLikeFile(fileObj)) {
            SC_setPdfCropPreference(cropMode);
        }
        placementSetImportPageNumber(pageNum);

        var measureItem = null;
        try {
            measureItem = targetDoc.placedItems.add();
            measureItem.file = fileObj;
            return {
                width: measureItem.width,
                height: measureItem.height
            };
        } finally {
            if (measureItem) {
                try { measureItem.remove(); } catch (e) { }
            }
        }
    }

    /**
     * 各ページを一度ずつ一時配置して、素のサイズを計測する
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {File} fileObj - 読み込み元ファイル
     * @param {Array} targetPages - 対象ページ番号の配列
     * @param {number} cropMode - トリミング指定（CropTo の値）
     * @returns {Array} 計測結果の配列。計測できなかったページは null
     */
    function placementMeasurePages(targetDoc, fileObj, targetPages, cropMode) {
        var measured = [];
        for (var i = 0; i < targetPages.length; i++) {
            var pageNum = parseInt(targetPages[i], 10);
            if (isNaN(pageNum) || pageNum < 1) pageNum = 1;
            try {
                var size = placementMeasurePlacedPageSize(targetDoc, fileObj, pageNum, cropMode);
                measured.push({ page: pageNum, width: size.width, height: size.height });
            } catch (e) {
                measured.push(null); // 計測できないページ / Page that cannot be measured
            }
        }
        return measured;
    }

    /**
     * 計測結果から各ページの配置位置と全体サイズを求める（DOM は参照しない）
     * @param {Array} measured - placementMeasurePages() の戻り値
     * @param {number} scaleFactor - 配置倍率（1 = 100%）
     * @param {number} gap - ページ間の間隔（pt）
     * @param {number} colsPerRow - 1 行あたりの列数。0 は自動
     * @returns {object} slots（配置位置の配列）と width / height を持つオブジェクト
     */
    function placementBuildLayout(measured, scaleFactor, gap, colsPerRow) {
        var slots = [];
        var nextX = 0;
        var topY = 0;
        var colCount = 0;
        var rowMaxH = 0; // 現在の行で最も高いページの高さ / Tallest page height in the current row
        var maxRight = 0;
        var minBottom = 0;

        for (var i = 0; i < measured.length; i++) {
            var m = measured[i];
            if (!m) {
                slots.push(null);
                continue;
            }

            var pageW = m.width * scaleFactor;
            var pageH = m.height * scaleFactor;

            // 指定列数に達した、または行幅が上限を超えるとき次の行へ折り返す
            // Wrap to the next row when the column count or the row width limit is reached
            var wrapByCol = colsPerRow > 0 && colCount >= colsPerRow;
            var wrapByEdge = colsPerRow === 0 && nextX > 0 && nextX + pageW > MAX_AUTO_ROW_WIDTH;
            if (wrapByCol || wrapByEdge) {
                topY -= rowMaxH + gap;
                nextX = 0;
                rowMaxH = 0;
                colCount = 0;
            }

            slots.push({ page: m.page, x: nextX, y: topY, width: pageW, height: pageH });

            if (nextX + pageW > maxRight) maxRight = nextX + pageW;
            if (topY - pageH < minBottom) minBottom = topY - pageH;

            nextX += pageW + gap;
            colCount++;
            if (pageH > rowMaxH) rowMaxH = pageH;
        }

        return { slots: slots, width: maxRight, height: -minBottom };
    }

    /**
     * 1 枚目はアクティブアートボードを再利用し、以降は新規追加する
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {number} activeIdx - アクティブアートボードのインデックス
     * @param {Array} abRect - アートボードの矩形 [left, top, right, bottom]
     * @param {number} abCount - これまでに用意したアートボード数
     * @returns {number} 更新後のアートボード数
     */
    function placementUseOrAddArtboard(targetDoc, activeIdx, abRect, abCount) {
        if (abCount === 0) {
            targetDoc.artboards[activeIdx].artboardRect = abRect;
        } else {
            targetDoc.artboards.add(abRect);
        }
        return abCount + 1;
    }

    /**
     * ページを配置する（必要なら倍率を適用）
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {File} fileObj - 読み込み元ファイル
     * @param {number} pageNum - 配置するページ番号
     * @param {Array} pos - 配置位置 [left, top]
     * @param {number} cropMode - トリミング指定（CropTo の値）
     * @param {number} scalePct - 配置倍率（%）
     * @returns {PlacedItem} 配置したアイテム
     */
    function placementPlacePage(targetDoc, fileObj, pageNum, pos, cropMode, scalePct) {
        if (isPdfLikeFile(fileObj)) {
            SC_setPdfCropPreference(cropMode);
        }
        placementSetImportPageNumber(pageNum);

        var item = targetDoc.placedItems.add();
        item.file = fileObj;

        if (typeof scalePct === "number" && scalePct !== 100) {
            item.resize(scalePct, scalePct);
        }
        // resize() の基準点は position と無関係なため、スケール後に左上位置を確定させる
        // resize() uses its own anchor, so set the top-left position after scaling
        item.position = pos;
        return item;
    }

    /**
     * 矩形群 [left, top, right, bottom] の和を求める
     * @param {Array} rects - 矩形の配列
     * @returns {Array} 外接矩形。求められない場合は null
     */
    function placementUnionRects(rects) {
        var union = null;
        for (var i = 0; i < rects.length; i++) {
            var r = rects[i];
            if (!r || r.length !== 4) continue;
            if (!union) {
                union = [r[0], r[1], r[2], r[3]];
                continue;
            }
            if (r[0] < union[0]) union[0] = r[0];
            if (r[1] > union[1]) union[1] = r[1];
            if (r[2] > union[2]) union[2] = r[2];
            if (r[3] < union[3]) union[3] = r[3];
        }
        return union;
    }

    /**
     * 全アートボードの外接範囲を求める
     * @param {Document} targetDoc - 対象ドキュメント
     * @returns {Array} 外接矩形。求められない場合は null
     */
    function placementGetArtboardsBounds(targetDoc) {
        if (!targetDoc || !targetDoc.artboards) return null;
        var rects = [];
        for (var i = 0; i < targetDoc.artboards.length; i++) {
            rects.push(targetDoc.artboards[i].artboardRect);
        }
        return placementUnionRects(rects);
    }

    /**
     * 複数アイテムの可視境界の和を求める
     * @param {Array} items - 対象アイテムの配列
     * @returns {Array} 外接矩形。求められない場合は null
     */
    function placementGetItemsVisibleBounds(items) {
        if (!items) return null;
        var rects = [];
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (!it) continue;
            try {
                rects.push(it.visibleBounds);
            } catch (e) {
                try { rects.push(it.geometricBounds); } catch (err) { }
            }
        }
        return placementUnionRects(rects);
    }

    /**
     * 指定した範囲全体が見えるよう表示倍率と表示位置を調整する
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {Array} bounds - 表示したい矩形 [left, top, right, bottom]
     * @returns {void}
     */
    function placementFitBoundsInView(targetDoc, bounds) {
        if (!targetDoc || !bounds) return;

        try {
            app.activeDocument = targetDoc;
        } catch (e) { }

        try {
            var view = targetDoc.activeView;
            if (!view) return;

            var unionWidth = bounds[2] - bounds[0];
            var unionHeight = bounds[1] - bounds[3];
            if (unionWidth <= 0 || unionHeight <= 0) return;

            var currentBounds = view.bounds;
            var currentWidth = currentBounds[2] - currentBounds[0];
            var currentHeight = currentBounds[1] - currentBounds[3];
            var currentZoom = view.zoom;
            if (!(currentWidth > 0) || !(currentHeight > 0) || !(currentZoom > 0)) return;

            var paddingScale = 0.9;
            var zoomX = currentZoom * (currentWidth / unionWidth);
            var zoomY = currentZoom * (currentHeight / unionHeight);

            view.centerPoint = [
                (bounds[0] + bounds[2]) / 2,
                (bounds[1] + bounds[3]) / 2
            ];
            view.zoom = Math.min(zoomX, zoomY) * paddingScale;
        } catch (e) { }
    }

    /**
     * 最大キャンバス範囲を求める（OMOTI 氏のアイデア）
     * @param {Document} targetDoc - 対象ドキュメント
     * @returns {Array} キャンバスの矩形 [left, top, right, bottom]
     */
    function getLargestCanvasBounds(targetDoc) {
        var LARGEST_SIZE = 16383;
        try {
            var tempLayer = targetDoc.layers.add();
            var tempText = tempLayer.textFrames.add();
            var left = tempText.matrix.mValueTX;
            var top = tempText.matrix.mValueTY;
            tempLayer.remove();
            return [left, top, left + LARGEST_SIZE, top - LARGEST_SIZE];
        } catch (e) {
            var half = LARGEST_SIZE / 2;
            return [-half, half, half, -half];
        }
    }

    // =========================================
    // ケイ処理ヘルパー
    // =========================================

    /**
     * 配置アイテムの外接矩形でクリッピングマスクグループを作成する
     * @param {PlacedItem} placedItem - 対象の配置アイテム
     * @returns {GroupItem} 作成したクリッピンググループ
     */
    function keiCreateClippingMaskGroup(placedItem) {
        var targetLayer = placedItem.layer;
        var rect = targetLayer.pathItems.rectangle(
            placedItem.top,
            placedItem.left,
            placedItem.width,
            placedItem.height
        );
        rect.stroked = false;
        rect.filled = false;

        var groupItem = targetLayer.groupItems.add();
        placedItem.moveToBeginning(groupItem);
        rect.moveToBeginning(groupItem);
        groupItem.clipped = true;

        return groupItem;
    }

    /**
     * 角丸 LiveEffect を適用する
     * @param {PageItem} targetItem - 対象アイテム
     * @param {number} radius - 角丸の半径（pt）
     * @returns {void}
     */
    function keiApplyRoundCornersLiveEffect(targetItem, radius) {
        if (!targetItem) return;
        var r = Number(radius);
        if (isNaN(r) || r <= 0) return;
        var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + r + ' "/></LiveEffect>';
        try {
            targetItem.applyEffect(xml);
        } catch (e) { }
    }

    /**
     * 配置アイテムにケイ処理を適用する
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {PlacedItem} placedItem - 対象の配置アイテム
     * @param {object} keiOpts - mode / roundCorners / roundRadius を持つ設定
     * @returns {PageItem} 処理後のアイテム。処理しない場合は元のアイテム
     */
    function keiApplyToPlacedItem(targetDoc, placedItem, keiOpts) {
        if (!keiOpts || !placedItem) return placedItem;

        if (keiOpts.mode === 'clipGroup') {
            var group = keiCreateClippingMaskGroup(placedItem);

            if (keiOpts.roundCorners && keiOpts.roundRadius > 0) {
                keiApplyRoundCornersLiveEffect(group, keiOpts.roundRadius);
            }

            targetDoc.selection = [group];
            app.executeMenuCommand('Adobe New Stroke Shortcut');
            app.executeMenuCommand('Live Pathfinder Exclude');
            targetDoc.selection = null;
            return group;
        }

        return placedItem;
    }

    // =========================================
    // 配置時のトリミング設定
    // 効くのは plugin/PDFImport/CropTo。値は 0=アート / 1=トリミング（CropBox）/ 2=仕上がり（TrimBox）/ 3=裁ち落とし / 4=メディア（実測）
    // =========================================

    // 配置時に使うトリミング。「トリミング」＝ CropBox 固定
    // Crop box used when placing; fixed to CropBox
    var DEFAULT_CROP_MODE = 1;

    // PDF 配置時の crop プリファレンスキー（環境差を吸収するため複数試行）/ Crop preference keys (multiple keys tried for version differences)
    var PDF_CROP_PREFERENCE_KEYS = [
        "plugin/PDFImport/CropToBox",
        "plugin/PDFImport/CropTo",
        "plugin/PDFImport/CropBox",
        "plugin/PDFImport/CropToType"
    ];

    /**
     * crop 種別を複数キーに設定する（環境差を吸収）
     * @param {number} cropVal - 設定する crop 種別
     * @returns {void}
     */
    function SC_setPdfCropPreference(cropVal) {
        for (var i = 0; i < PDF_CROP_PREFERENCE_KEYS.length; i++) {
            try {
                app.preferences.setIntegerPreference(PDF_CROP_PREFERENCE_KEYS[i], cropVal);
            } catch (e) { }
        }
    }

    /**
     * 現在の crop プリファレンスを退避する
     * @returns {object} キーと値を持つスナップショット
     */
    function SC_snapshotPdfCropPreference() {
        var snapshot = {};
        for (var i = 0; i < PDF_CROP_PREFERENCE_KEYS.length; i++) {
            try {
                snapshot[PDF_CROP_PREFERENCE_KEYS[i]] = app.preferences.getIntegerPreference(PDF_CROP_PREFERENCE_KEYS[i]);
            } catch (e) { }
        }
        return snapshot;
    }

    /**
     * 退避した crop プリファレンスを復元する
     * @param {object} snapshot - SC_snapshotPdfCropPreference() の戻り値
     * @returns {void}
     */
    function SC_restorePdfCropPreference(snapshot) {
        if (!snapshot) return;
        for (var i = 0; i < PDF_CROP_PREFERENCE_KEYS.length; i++) {
            var prefKey = PDF_CROP_PREFERENCE_KEYS[i];
            if (snapshot[prefKey] === undefined) continue;
            try {
                app.preferences.setIntegerPreference(prefKey, snapshot[prefKey]);
            } catch (e) { }
        }
    }

    // ============================================================
    // ページ数取得ヘルパー
    // - リンクされた PDF/AI の総ページ数（最終ページ番号）を推定
    // - ファイルが明示指定された場合は、一時配置して既存の選択ベースの
    //   ページ数取得ロジックを再利用し、取得後すぐに削除
    // - 現在の選択に対してリンク変更や内容変更は行わない
    // ============================================================

    /**
     * 拡張子から PDF/AI ファイルかどうかを判定する
     * @param {File} f - 対象ファイル
     * @returns {boolean} PDF または AI なら true
     */
    function isPdfLikeFile(f) {
        return /\.(?:pdf|ai)$/i.test(String((f && f.name) || ""));
    }

    /**
     * 選択範囲から最初の PlacedItem を再帰的に探す
     * @param {Array} items - 探索対象のアイテム配列
     * @returns {PlacedItem} 見つかった配置アイテム。無ければ null
     */
    function pageCountFindFirstPlacedItem(items) {
        if (!items || items.length <= 0) return null;
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) continue;
            var typeName = (item.constructor && item.constructor.name) ? item.constructor.name : '';
            if (typeName === 'PlacedItem') return item;
            if (typeName === 'GroupItem') {
                var hit = pageCountFindFirstPlacedItem(item.pageItems);
                if (hit) return hit;
            }
        }
        return null;
    }

    /**
     * PDF/AI ファイルから総ページ数を推定する
     * @param {File} file - 対象ファイル
     * @returns {number} 推定した総ページ数。取得できない場合は undefined
     */
    function getPageLengthFromFile(file) {
        var countInDictPattern = /<<\/Count\s(\d+)/;
        var pageRefPattern1 = /<<\/Type\/Page\/Parent/;
        var pageRefPattern2 = /\/Type\s\/Page\s/;
        var pageRefPattern3 = /\/StructParents\s\d+.*\/Type\/Page>>/;
        var linearizedCountPattern = /<<\/Linearized\s.+\/N\s(\d+)\/T\s.+>>/;
        var pagesNodePattern = /\/Type\/Pages/;
        var countPattern = /\/Count\s(\d+)/;

        var pageCount;
        var line;
        var pageRefCount = 0;
        var declaredCount = 0;
        var hasDeclaredCount = false;

        try {
            file.open('r');
            while (!file.eof) {
                line = file.readln();
                if (countInDictPattern.test(line) || linearizedCountPattern.test(line)) {
                    declaredCount = Number(RegExp.$1);
                    if (declaredCount > 0) {
                        pageCount = declaredCount;
                        hasDeclaredCount = true;
                        break;
                    }
                }
                if (pageRefPattern1.test(line) || pageRefPattern2.test(line) || pageRefPattern3.test(line)) pageRefCount++;
                if (pagesNodePattern.test(line)) {
                    line = file.readln();
                    if (countPattern.test(line)) {
                        declaredCount = Number(RegExp.$1);
                        if (pageRefCount < declaredCount) pageRefCount = declaredCount;
                    }
                }
            }
            // 宣言値を取得できたときは、数え上げた概算で上書きしない
            // Keep the declared count; do not overwrite it with the scanned estimate
            if (!hasDeclaredCount && pageRefCount > 0) pageCount = pageRefCount;
        } catch (e) {
            SC_alert(LABELS.alert.pageCountFail, e);
        } finally {
            try { file.close(); } catch (err) { }
        }
        return pageCount;
    }

    /**
     * 選択中の配置画像のリンク元から総ページ数を取得する
     * @param {Array} selectionItems - 対象アイテムの配列
     * @returns {number} 総ページ数。取得できない場合は null
     */
    function getLastPageFromSelection(selectionItems) {
        var placed = pageCountFindFirstPlacedItem(selectionItems);
        if (!placed) return null;

        var f = placed.file;
        if (!f) {
            SC_alert(LABELS.alert.linkUnknown);
            return null;
        }

        if (!isPdfLikeFile(f)) return null;

        var last = getPageLengthFromFile(f);
        if (!last || isNaN(Number(last)) || Number(last) <= 0) {
            SC_alert(LABELS.alert.pageCountFail);
            return null;
        }

        return Number(last);
    }

    /**
     * 指定ファイル（または現在の選択）から総ページ数を取得する
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {File} fileObjOrNull - 読み込み元ファイル。null なら選択から取得
     * @param {function} setPathTextFn - ファイルパス表示を更新するコールバック
     * @returns {number} 総ページ数。取得できない場合は null
     */
    function updatePageCountFromPlacedOrFile(targetDoc, fileObjOrNull, setPathTextFn) {
        var placedTemp = null;

        try {
            // ファイル指定時は一時配置して、既存の選択ベースのページ数取得ロジックを再利用
            if (fileObjOrNull) {
                if (!isPdfLikeFile(fileObjOrNull)) {
                    SC_alert(LABELS.alert.pickPdfAi);
                    return null;
                }

                try {
                    placedTemp = targetDoc.placedItems.add();
                    placedTemp.file = fileObjOrNull;

                    // 一時配置物は、いったん表示中ビューの左上付近へ置く
                    try {
                        var vb = targetDoc.activeView && targetDoc.activeView.bounds ? targetDoc.activeView.bounds : null;
                        if (vb && vb.length === 4) {
                            placedTemp.position = [vb[0], vb[1]];
                        }
                    } catch (e) { }

                    // 既存ロジックをそのまま使って総ページ数を取得
                    var lastFromFile = getLastPageFromSelection([placedTemp]);
                    if (setPathTextFn) setPathTextFn(fileObjOrNull);
                    return lastFromFile;
                } finally {
                    if (placedTemp) {
                        try { placedTemp.remove(); } catch (e) { }
                    }
                }
            }

            // ファイル未指定時は現在の選択から取得（PlacedItem が選択されている前提）
            var last = getLastPageFromSelection(targetDoc.selection);

            // 選択中の配置画像があれば、そのファイルパス表示も更新
            try {
                var placedSel = pageCountFindFirstPlacedItem(targetDoc.selection);
                if (placedSel && placedSel.file && setPathTextFn) setPathTextFn(placedSel.file);
            } catch (e) { }

            return last;
        } catch (e) {
            SC_alert(LABELS.alert.pageCountFail, e);
            return null;
        }
    }

    /**
     * 入力文字列（例: "1-20", "1,3,5"）をページ番号配列へ変換する
     * 重複と範囲外（1..maxPages 以外）は除外する
     * @param {string} inputStr - ページ範囲の文字列
     * @param {number} maxPages - 上限ページ数。0 以下なら上限なし
     * @returns {Array} ページ番号の配列
     */
    function parsePageNumbers(inputStr, maxPages) {
        var result = [];
        var seen = {};
        var hasMax = (typeof maxPages === "number" && maxPages > 0);

        /**
         * 1 件のページ番号を検証して追加する
         * @param {number} num - ページ番号
         * @returns {void}
         */
        function addPage(num) {
            if (isNaN(num) || num < 1) return;
            if (hasMax && num > maxPages) return;
            var key = "p" + num;
            if (seen[key]) return;
            seen[key] = true;
            result.push(num);
        }

        var parts = String(inputStr).split(',');
        for (var i = 0; i < parts.length; i++) {
            var part = parts[i].replace(/^\s+|\s+$/g, '');
            if (part === '') continue;

            if (part.indexOf('-') > -1) {
                var bounds = part.split('-');
                var start = parseInt(bounds[0], 10);
                var end = parseInt(bounds[1], 10);
                if (!isNaN(start) && !isNaN(end)) {
                    var min = Math.min(start, end);
                    var max = Math.max(start, end);
                    for (var j = min; j <= max; j++) {
                        addPage(j);
                    }
                }
            } else {
                addPage(parseInt(part, 10));
            }
        }
        return result;
    }

    /**
     * メイン処理。ダイアログを構築して配置を実行する
     * @returns {void}
     */
    function main() {

        if (app.documents.length === 0) {
            SC_alert(LABELS.alert.needDoc);
            return;
        }

        var doc = app.activeDocument;
        // 現在の読み込み対象ファイル。未選択時は null。
        var sourceFile = null;
        var detectedRangeText = "";
        var customRangeText = "";
        // 直前に［指定ページ］が選択されていたか / Whether custom range was the previous mode
        var wasCustomRange = false;
        // OK 押下時の配置指示。ダイアログを閉じてから実行する / Placement request, run after the dialog closes
        var placementRequest = null;
        // 進捗表示用のパレット / Progress palette
        var progressWin = null;
        var progressBar = null;

        // ------------------------
        // UI構築
        // ------------------------
        var win = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        win.alignChildren = "fill";

        var bodyGroup = win.add("group");
        bodyGroup.orientation = "row";
        bodyGroup.alignChildren = ["fill", "top"];

        var leftColumnGroup = bodyGroup.add("group");
        leftColumnGroup.orientation = "column";
        leftColumnGroup.alignChildren = "fill";

        var rightColumnGroup = bodyGroup.add("group");
        rightColumnGroup.orientation = "column";
        rightColumnGroup.alignChildren = "fill";

        var sourcePanel = leftColumnGroup.add('panel', undefined, getLabel(LABELS.panel.source));
        setupPanel(sourcePanel);
        var btnBrowse = sourcePanel.add('button', undefined, getLabel(LABELS.button.selectFile));
        // パネルの alignChildren は fill なので、ボタンだけ自前の幅で左寄せにする
        // The panel fills its children, so keep the button at its natural width
        btnBrowse.alignment = ["left", "center"];
        btnBrowse.helpTip = getLabel(LABELS.tooltip.source);
        var stSourceName = sourcePanel.add('statictext', undefined, getLabel(LABELS.label.notSelected));
        stSourceName.characters = 16;

        var pagesPanel = leftColumnGroup.add('panel', undefined, getLabel(LABELS.panel.pages));
        setupPanel(pagesPanel);
        var rangeModeGroup = pagesPanel.add('group');
        setupGroup(rangeModeGroup, 'column');
        var rbRangeAll = rangeModeGroup.add('radiobutton', undefined, getLabel(LABELS.radio.rangeAll));
        var rbRangeFirst = rangeModeGroup.add('radiobutton', undefined, getLabel(LABELS.radio.rangeFirst));
        var rbRangeCustom = rangeModeGroup.add('radiobutton', undefined, labelText(LABELS.radio.rangeCustom));
        // 初期状態は「全ページ」を選択
        rbRangeAll.value = true;

        var rangeInputGroup = pagesPanel.add('group');
        setupGroup(rangeInputGroup, 'row');
        var etRange = rangeInputGroup.add('edittext', undefined, '');
        etRange.characters = 10;
        // 初期選択は「全ページ」なので無効から始める / Starts disabled: the default mode is All Pages
        etRange.enabled = false;
        etRange.helpTip = getLabel(LABELS.tooltip.customRange);

        var totalPagesGroup = pagesPanel.add('group');
        totalPagesGroup.orientation = 'row';
        totalPagesGroup.alignChildren = ['right', 'center'];
        totalPagesGroup.alignment = ['fill', 'top'];
        var stTotalPages = totalPagesGroup.add('statictext', undefined, '');
        stTotalPages.characters = 8;
        stTotalPages.justify = 'right';
        stTotalPages.alignment = ['right', 'center'];
        stTotalPages.helpTip = getLabel(LABELS.tooltip.totalPages);

        var methodPanel = leftColumnGroup.add("panel", undefined, getLabel(LABELS.panel.mode));
        setupPanel(methodPanel);
        var methodGroup = methodPanel.add("group");
        setupGroup(methodGroup, "column");
        var rbPerArtboard = methodGroup.add("radiobutton", undefined, getLabel(LABELS.radio.perArtboard));
        rbPerArtboard.helpTip = getLabel(LABELS.tooltip.perArtboard);
        var rbIgnoreArtboard = methodGroup.add("radiobutton", undefined, getLabel(LABELS.radio.ignoreArtboard));
        rbIgnoreArtboard.helpTip = getLabel(LABELS.tooltip.ignoreArtboard);
        rbPerArtboard.value = true;

        var layoutPanel = rightColumnGroup.add("panel", undefined, getLabel(LABELS.panel.placement));
        setupPanel(layoutPanel);

        var columnsGroup = layoutPanel.add("group");
        setupGroup(columnsGroup, "row");
        var stColsLabel = columnsGroup.add("statictext", undefined, labelText(LABELS.label.columns));
        stColsLabel.justify = "right";
        var etCols = columnsGroup.add("edittext", undefined, getLabel(LABELS.label.columnsAuto));
        etCols.characters = 5;
        etCols.helpTip = getLabel(LABELS.tooltip.columns);
        stColsLabel.helpTip = getLabel(LABELS.tooltip.columns);

        var rowsGroup = layoutPanel.add("group");
        setupGroup(rowsGroup, "row");
        var stRowsLabel = rowsGroup.add("statictext", undefined, labelText(LABELS.label.rows));
        stRowsLabel.justify = "right";
        var stRowsValue = rowsGroup.add("statictext", undefined, "");
        stRowsValue.characters = 5;
        stRowsLabel.helpTip = getLabel(LABELS.tooltip.estimate);
        stRowsValue.helpTip = getLabel(LABELS.tooltip.estimate);

        var gapGroup = layoutPanel.add("group");
        setupGroup(gapGroup, "row");
        var stGapLabel = gapGroup.add("statictext", undefined, labelText(LABELS.label.gap));
        stGapLabel.justify = "right";
        var etArtboardGap = gapGroup.add("edittext", undefined, String(DEFAULT_ARTBOARD_GAP));
        etArtboardGap.characters = 5;
        var stGapUnit = gapGroup.add("statictext", undefined, getLabel(LABELS.label.gapUnit));

        // 倍率は［アートボードを無視］のときだけ効く条件付き項目なので、レイアウトから分ける
        // Scale only applies when ignoring artboards, so keep it out of the Layout panel
        var optionPanel = rightColumnGroup.add("panel", undefined, getLabel(LABELS.panel.option));
        setupPanel(optionPanel);

        var scaleGroup = optionPanel.add("group");
        setupGroup(scaleGroup, "row");
        var stScaleLabel = scaleGroup.add("statictext", undefined, labelText(LABELS.label.scale));
        stScaleLabel.justify = "right";
        var etScale = scaleGroup.add("edittext", undefined, "100");
        etScale.characters = 5;
        etScale.helpTip = getLabel(LABELS.tooltip.scale);
        stScaleLabel.helpTip = getLabel(LABELS.tooltip.scale);
        var stScaleUnit = scaleGroup.add("statictext", undefined, getLabel(LABELS.label.scaleUnit));

        var keiPanel = rightColumnGroup.add("panel", undefined, getLabel(LABELS.panel.kei));
        setupPanel(keiPanel);
        var keiModeGroup = keiPanel.add("group");
        setupGroup(keiModeGroup, "column");
        var rbKeiNone = keiModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.keiNone));
        var rbKeiClipGroup = keiModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.keiClipGroup));
        rbKeiNone.value = true;

        var roundCornerGroup = keiPanel.add("group");
        setupGroup(roundCornerGroup, "row");
        var cbRoundCorner = roundCornerGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.roundCorner));
        var etRoundCorner = roundCornerGroup.add("edittext", undefined, "3");
        etRoundCorner.characters = 5;
        var stRoundCornerUnit = roundCornerGroup.add("statictext", undefined, getUnitInfo().label);
        cbRoundCorner.value = false;
        etRoundCorner.enabled = false;
        cbRoundCorner.helpTip = getLabel(LABELS.tooltip.roundCorner);
        etRoundCorner.helpTip = getLabel(LABELS.tooltip.roundCorner);

        // === ボタンエリア（左スペーサー／右キャンセル・OK）/ Button area (spacer left, cancel+ok right)
        var btnRowGroup = win.add("group");
        btnRowGroup.alignment = ["fill", "top"];
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["fill", "center"];
        btnRowGroup.margins = [0, 5, 0, 0]; // ボタンエリア上マージン +5 / extra top margin
        btnRowGroup.spacing = 0;

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "center"];
        spacer.minimumSize.width = 0;
        spacer.maximumSize.height = 0;

        // 右グループ（キャンセル・OKボタン）/ Right group (Cancel/OK buttons)
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.orientation = "row";
        btnRightGroup.alignment = ["right", "center"];
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = 10;
        var btnCancel = btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOk = btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        btnOk.alignment = ["right", "center"];

        // ------------------------
        // UI状態の更新
        // ------------------------

        /**
         * 読み込みファイルの有無に応じて各パネルの有効／無効を切り替える
         * @returns {void}
         */
        function updatePanelEnabledState() {
            var enabled = !!sourceFile;
            pagesPanel.enabled = enabled;
            methodPanel.enabled = enabled;
            layoutPanel.enabled = enabled;
            optionPanel.enabled = enabled;
            keiPanel.enabled = enabled;
        }

        /**
         * ページ範囲入力欄の有効／無効と表示内容を切り替える
         * @returns {void}
         */
        function updateRangeEnabledState() {
            if (rbRangeCustom.value) {
                etRange.enabled = true;
                etRange.text = customRangeText || detectedRangeText || '';
            } else {
                // ［指定ページ］から離れるときだけ入力内容を退避する
                // Keep the typed range only when leaving custom mode
                if (wasCustomRange) customRangeText = etRange.text;
                etRange.enabled = false;
                etRange.text = '';
            }
            wasCustomRange = rbRangeCustom.value;
        }

        /**
         * 検出済みの総ページ数を取得する
         * @returns {number} 総ページ数。未検出なら 0
         */
        function getDetectedTotalPages() {
            if (!detectedRangeText) return 0;
            var m = detectedRangeText.match(/^(?:1-)?(\d+)$/);
            return m ? (parseInt(m[1], 10) || 0) : 0;
        }

        /**
         * 現在の設定で配置されるページ数を求める
         * @param {number} totalPages - 総ページ数
         * @returns {number} 配置されるページ数
         */
        function getPlacedPageCount(totalPages) {
            if (rbRangeAll.value) return totalPages;
            if (rbRangeFirst.value) return totalPages > 0 ? 1 : 0;
            return parsePageNumbers(customRangeText || detectedRangeText || '', totalPages).length;
        }

        /**
         * 「配置ページ数 / 総ページ数」の表示を更新する
         * @returns {void}
         */
        function updateTotalPagesLabel() {
            var totalPages = getDetectedTotalPages();
            var placedPages = getPlacedPageCount(totalPages);
            stTotalPages.text = totalPages > 0 ? (placedPages + ' / ' + totalPages) : '';
        }

        /**
         * 行数の表示を更新する
         * 列数が自動のときは折り返し位置がページ幅次第で確定しないため、空欄にしてディム表示にする
         * @returns {void}
         */
        function updateRowsInfo() {
            var colsPerRow = getColsPerRow();
            var placedPages = getPlacedPageCount(getDetectedTotalPages());
            var known = colsPerRow > 0;

            stRowsLabel.enabled = known;
            stRowsValue.enabled = known;
            stRowsValue.text = (known && placedPages > 0) ? String(Math.ceil(placedPages / colsPerRow)) : '';
        }

        /**
         * 倍率欄の有効／無効を切り替える
         * @returns {void}
         */
        function updateScaleEnabledState() {
            var enabled = !rbPerArtboard.value;
            etScale.enabled = enabled;
            stScaleLabel.enabled = enabled;
            stScaleUnit.enabled = enabled;
        }

        /**
         * 配置方法に合わせて間隔欄のツールチップを切り替える
         * @returns {void}
         */
        function updateGapHelpTip() {
            var helpText = rbPerArtboard.value ? getLabel(LABELS.tooltip.gapPerArtboard) : getLabel(LABELS.tooltip.gapIgnoreArtboard);
            stGapLabel.helpTip = helpText;
            etArtboardGap.helpTip = helpText;
            stGapUnit.helpTip = helpText;
        }

        // ------------------------
        // UI表示補助
        // ------------------------

        /**
         * ファイル表示用の名前とパスを求める
         * @param {File} f - 対象ファイル
         * @returns {object} name と path を持つオブジェクト
         */
        function getFileDisplayInfo(f) {
            if (!f) {
                return {
                    name: getLabel(LABELS.label.notSelected),
                    path: ''
                };
            }

            var nameText;
            var pathText;

            try {
                nameText = decodeURIComponent(f.name);
            } catch (e) {
                nameText = String(f.name || getLabel(LABELS.label.notSelected));
            }

            try {
                pathText = decodeURIComponent(f.fsName);
            } catch (e) {
                pathText = String(f.fsName || '');
            }

            return {
                name: nameText,
                path: pathText
            };
        }

        /**
         * File.openDialog のフィルター。フォルダーは辿れるよう通す
         * @param {File} fileObj - 判定対象
         * @returns {boolean} 表示するなら true
         */
        function isPdfOrAiFile(fileObj) {
            if (!fileObj) return false;
            if (fileObj instanceof Folder) return true;
            return isPdfLikeFile(fileObj);
        }

        /**
         * 読み込み対象ファイルを設定し、名前とパスの表示を更新する
         * @param {File} f - 読み込み元ファイル。null で未指定に戻す
         * @returns {void}
         */
        function setPathText(f) {
            sourceFile = f || null;
            updatePanelEnabledState();

            var info = getFileDisplayInfo(f);
            stSourceName.text = info.name;
            stSourceName.helpTip = info.path;
        }

        /**
         * 数値欄に ↑↓ キーでの増減を割り当てる
         * @param {object} editText - 対象の edittext
         * @returns {void}
         */
        function changeValueByArrowKey(editText) {
            editText.addEventListener("keydown", function (event) {
                var textValue = String(editText.text);
                var value;

                if (textValue === getLabel(LABELS.label.columnsAuto)) {
                    if (event.keyName == "Up") {
                        event.preventDefault();
                        editText.text = 1;
                        try {
                            if (typeof editText.onChange === "function") {
                                editText.onChange();
                            }
                        } catch (e) { }
                    }
                    return;
                }

                value = Number(textValue);
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

                value = Math.round(value);
                editText.text = value;

                try {
                    if (typeof editText.onChange === "function") {
                        editText.onChange();
                    }
                } catch (e) { }
            });
        }

        // ------------------------
        // 配置オプション取得
        // ------------------------

        /**
         * UI で指定した 1 行あたりの列数を取得する
         * @returns {number} 列数。自動のときは 0
         */
        function getColsPerRow() {
            var v = parseInt(etCols.text, 10);
            if (isNaN(v) || v <= 0) return 0;
            return v;
        }

        /**
         * UI で指定したページ間隔を取得する
         * @returns {number} 間隔（pt）
         */
        function getArtboardGap() {
            var v = parseFloat(etArtboardGap.text);
            if (isNaN(v) || v < 0) v = DEFAULT_ARTBOARD_GAP;
            return v;
        }

        /**
         * UI で指定した配置倍率を取得する
         * @returns {number} 配置倍率（%）
         */
        function getScaleFromUI() {
            var v = parseFloat(etScale.text);
            if (isNaN(v) || v <= 0) v = 100;
            return v;
        }

        /**
         * UI で指定したケイ処理の設定を取得する
         * @returns {object} ケイ処理の設定。処理しない場合は null
         */
        function getKeiOptionsFromUI() {
            if (rbKeiNone.value) {
                return null;
            }
            var radius = parseFloat(etRoundCorner.text);
            if (isNaN(radius) || radius < 0) radius = 0;
            return {
                mode: 'clipGroup',
                roundCorners: cbRoundCorner.value,
                roundRadius: radius * getUnitInfo().pointsPerUnit
            };
        }

        /**
         * UI で指定した対象ページ番号の配列を取得する
         * @returns {Array} ページ番号の配列
         */
        function getTargetPagesFromUI() {
            if (rbRangeAll.value) {
                var lastPage = getDetectedTotalPages();
                if (!lastPage || lastPage < 1) lastPage = 1;
                var pages = [];
                for (var p = 1; p <= lastPage; p++) {
                    pages.push(p);
                }
                return pages;
            }
            if (rbRangeFirst.value) {
                return [1];
            }
            var custom = parsePageNumbers(etRange.text, getDetectedTotalPages());
            return (custom && custom.length > 0) ? custom : [1];
        }

        // ------------------------
        // 進捗表示
        // ダイアログを閉じた後に使うため、独立したパレットで表示する
        // Shown in its own palette because the dialog is already closed
        // ------------------------

        /**
         * 進捗表示用のパレットを開く。開けなくても処理は継続する
         * @returns {void}
         */
        function showProgress() {
            try {
                progressWin = new Window("palette", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
                progressWin.alignChildren = "fill";
                progressBar = progressWin.add("progressbar", undefined, 0, 100);
                progressBar.preferredSize.width = 300;
                progressBar.preferredSize.height = 8;
                progressWin.show();
                progressWin.update();
            } catch (e) {
                progressWin = null;
                progressBar = null;
            }
        }

        /**
         * 進捗を更新する
         * @param {number} current - 処理済みの件数
         * @param {number} total - 総件数
         * @returns {void}
         */
        function updateProgress(current, total) {
            if (!progressWin || !progressBar || !total) return;
            try {
                progressBar.value = Math.round((current / total) * 100);
                progressWin.update();
            } catch (e) { }
        }

        /**
         * 進捗表示を閉じる
         * @returns {void}
         */
        function hideProgress() {
            if (!progressWin) return;
            try { progressWin.close(); } catch (e) { }
            progressWin = null;
            progressBar = null;
        }

        // ------------------------
        // 実行処理
        // ------------------------

        /**
         * 指定ページを配置する。アートボードごと／無視のどちらにも対応
         * @param {Array} targetPages - 対象ページ番号の配列
         * @param {object} keiOpts - ケイ処理の設定。null で処理なし
         * @param {boolean} perArtboard - true でページごとにアートボードを作成
         * @param {number} scale - 配置倍率（%）。perArtboard が true のときは無視
         * @returns {void}
         */
        function placePages(targetPages, keiOpts, perArtboard, scale) {
            var cropMode = DEFAULT_CROP_MODE;
            var scaleFactor = perArtboard ? 1 : (scale / 100);
            var gap = getArtboardGap();
            var colsPerRow = getColsPerRow();

            // crop 環境設定は計測で書き換わるため、計測より前に退避する
            // Snapshot the crop preferences before measuring, which overwrites them
            var cropSnapshot = SC_snapshotPdfCropPreference();
            var activeIdx = doc.artboards.getActiveArtboardIndex();
            var abCount = 0;
            var skippedCount = 0;
            var completed = false;
            var placedItems = [];
            showProgress();

            try {
                // 全ページを一度だけ計測し、レイアウト全体をカンバス中心へ揃える
                // Measure every page once, then center the whole layout on the canvas
                var measured = placementMeasurePages(doc, sourceFile, targetPages, cropMode);
                var layout = placementBuildLayout(measured, scaleFactor, gap, colsPerRow);
                var canvas = getLargestCanvasBounds(doc);
                var startX = (canvas[0] + canvas[2]) / 2 - layout.width / 2;
                var baseTop = (canvas[1] + canvas[3]) / 2 + layout.height / 2;

                for (var i = 0; i < layout.slots.length; i++) {
                    var slot = layout.slots[i];
                    if (!slot) {
                        skippedCount++; // 計測できなかったページ / Page that could not be measured
                    } else {
                        try {
                            var left = startX + slot.x;
                            var top = baseTop + slot.y;

                            if (perArtboard) {
                                abCount = placementUseOrAddArtboard(
                                    doc,
                                    activeIdx,
                                    [left, top, left + slot.width, top - slot.height],
                                    abCount
                                );
                            }

                            var item = placementPlacePage(doc, sourceFile, slot.page, [left, top], cropMode, perArtboard ? 100 : scale);
                            if (keiOpts) {
                                try { item = keiApplyToPlacedItem(doc, item, keiOpts); } catch (e) { }
                            }
                            placedItems.push(item);
                        } catch (e) {
                            skippedCount++; // 失敗ページはスキップして継続 / Skip the failed page and continue
                        }
                    }
                    updateProgress(i + 1, layout.slots.length);
                }
                completed = true;
            } catch (e) {
                SC_alert(LABELS.alert.placeError, e);
            } finally {
                try { placementResetImportPageNumber(); } catch (e) { }
                try { SC_restorePdfCropPreference(cropSnapshot); } catch (e) { }
                hideProgress();
            }

            if (!completed) return;

            if (perArtboard) {
                placementFitBoundsInView(doc, placementGetArtboardsBounds(doc));
            } else {
                // 配置したオブジェクトを選択して表示を合わせる / Select the placed objects and fit the view
                try { doc.selection = null; } catch (e) { }
                for (var j = 0; j < placedItems.length; j++) {
                    try { placedItems[j].selected = true; } catch (e) { }
                }
                placementFitBoundsInView(doc, placementGetItemsVisibleBounds(placedItems));
            }

            if (skippedCount > 0) SC_alert(LABELS.alert.someSkipped);
        }

        // ------------------------
        // UIイベント配線
        // ------------------------

        /**
         * ダイアログのイベントハンドラーを割り当てる
         * @returns {void}
         */
        function bindDialogEvents() {

            /**
             * ページ範囲まわりの表示をまとめて更新する
             * @returns {void}
             */
            function refreshRangeInfo() {
                updateRangeEnabledState();
                updateTotalPagesLabel();
                updateRowsInfo();
            }

            /**
             * 角丸まわりの有効／無効を切り替える
             * @returns {void}
             */
            function updateKeiRoundEnabled() {
                var keiActive = !rbKeiNone.value;
                cbRoundCorner.enabled = keiActive;
                etRoundCorner.enabled = keiActive && cbRoundCorner.value;
            }

            btnBrowse.onClick = function () {
                var f = File.openDialog(getLabel(LABELS.dialog.pickFile), isPdfOrAiFile);
                if (!f) return;
                var last = updatePageCountFromPlacedOrFile(doc, f, setPathText);
                detectedRangeText = last ? ('1-' + last) : '';
                if (!customRangeText) {
                    customRangeText = detectedRangeText;
                }
                refreshRangeInfo();
            };

            rbRangeAll.onClick = refreshRangeInfo;
            rbRangeFirst.onClick = refreshRangeInfo;
            rbRangeCustom.onClick = refreshRangeInfo;

            etRange.onChanging = function () {
                if (!rbRangeCustom.value) return;
                customRangeText = etRange.text;
                updateTotalPagesLabel();
                updateRowsInfo();
            };
            etRange.onChange = etRange.onChanging;

            etCols.onChanging = updateRowsInfo;
            etCols.onChange = function () {
                // 0 以下や数値でない入力は自動扱いなので、確定時に表示も「自動」へ揃える
                // Zero or non-numeric input means Auto; normalize the field text on commit
                var autoText = getLabel(LABELS.label.columnsAuto);
                if (getColsPerRow() === 0 && etCols.text !== autoText) etCols.text = autoText;
                updateRowsInfo();
            };

            rbPerArtboard.onClick = function () {
                updateScaleEnabledState();
                updateGapHelpTip();
                updateRowsInfo();
            };
            rbIgnoreArtboard.onClick = rbPerArtboard.onClick;

            rbKeiNone.onClick = updateKeiRoundEnabled;
            rbKeiClipGroup.onClick = updateKeiRoundEnabled;
            cbRoundCorner.onClick = updateKeiRoundEnabled;
            updateKeiRoundEnabled();

            btnCancel.onClick = function () {
                win.close(2);
            };

            btnOk.onClick = function () {
                if (!sourceFile) {
                    SC_alert(LABELS.alert.needFile);
                    return;
                }
                // 配置指示だけ控えて閉じる。実行は win.show() の後
                // Record the request and close; the placement runs after win.show()
                var perArtboard = rbPerArtboard.value;
                placementRequest = {
                    pages: getTargetPagesFromUI(),
                    keiOpts: getKeiOptionsFromUI(),
                    perArtboard: perArtboard,
                    scale: perArtboard ? 100 : getScaleFromUI()
                };
                win.close(1);
            };
        }

        // ------------------------
        // 初期状態反映
        // ------------------------

        /**
         * ダイアログの初期状態を反映する
         * @returns {void}
         */
        function initializeDialogState() {
            updatePanelEnabledState();

            var last = updatePageCountFromPlacedOrFile(doc, null, setPathText);
            detectedRangeText = last ? ('1-' + last) : '';
            if (!customRangeText) {
                customRangeText = detectedRangeText;
            }
            updateRangeEnabledState();
            updateTotalPagesLabel();
            updateScaleEnabledState();
            updateGapHelpTip();
            updateRowsInfo();

            stRoundCornerUnit.text = getUnitInfo().label;
            changeValueByArrowKey(etCols);
            changeValueByArrowKey(etArtboardGap);
            changeValueByArrowKey(etScale);
        }

        bindDialogEvents();
        initializeDialogState();

        if (win.show() !== 1 || !placementRequest) return;

        // モーダル表示中は executeMenuCommand が無視されることがあり、閉じ処理も詰まりやすい
        // Run after the dialog closes: executeMenuCommand can be ignored while a modal dialog is up
        placePages(
            placementRequest.pages,
            placementRequest.keiOpts,
            placementRequest.perArtboard,
            placementRequest.scale
        );

    }

    main();

}());
