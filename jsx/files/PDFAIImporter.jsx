#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

PDF/AI ファイルを指定したページ範囲で読み込み、現在のドキュメント上にページを配置します。
配置方法は「各ページを個別のアートボードとして並べる」「アートボードを追加せずオブジェクトとして配置する」から選べます。

### 対象ページ

- 全ページ（初期値）
- 先頭ページのみ
- 指定ページ（例: 1-10, 1,3,5）

読み込みファイルが確定するまでは、対象ページパネルと配置方法パネルは無効です。
選択中の配置画像、または［ファイル指定］で選んだ PDF/AI ファイルから総ページ数を推定し、
対象ページパネルに「実際に配置されるページ数／総ページ数」を表示します。
［ファイル指定］のダイアログでは PDF/AI ファイルのみを選択対象にします。

### 配置方法

- アートボードごと：各ページサイズに合わせてアートボードを作成／更新し、倍率 100% 固定で配置
- アートボードを無視：現在のアートボード左上を基準に、指定倍率でオブジェクトとして配置（配置後は対象を選択し、全体が見えるよう表示を調整）

### レイアウト

- 列数を指定すると、その数ごとに改行してグリッド状に配置
- 列数が「自動」のときは、カンバス右端に達したタイミングで折り返し
- 列数指定時のみ、列数入力欄の右側に推定の行 × 列を表示
- 間隔は配置方法で意味が変わる（アートボードごと＝アートボードの間隔／無視＝オブジェクト同士の間隔）

### ケイ

- なし
- ケイ線のみを追加（角丸オプションを有効にすると角丸半径を指定可能）

### 操作

- 列数・間隔・倍率の入力欄は ↑↓ で ±1、Shift+↑↓ で ±10
- 列数が「自動」のときは ↑ で 1 に切り替え
- 配置完了後は結果全体が見えるよう表示を自動調整（アートボードごと＝全アートボード表示／無視＝配置オブジェクト全体表示）

### 紹介記事

https://note.com/dtp_tranist/n/n42595650216f

更新日: 2026-04-13

*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.1.1";

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

    // 現在の UI 言語（ja / en）を判定 / Detect the current UI language (ja / en)
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
            kei: { ja: "枠線", en: "Stroke" }
        },
        radio: {
            rangeAll: { ja: "全ページ", en: "All Pages" },
            rangeFirst: { ja: "先頭ページのみ", en: "First Page Only" },
            rangeCustom: { ja: "指定ページ：", en: "Custom Pages:" },
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
            scale: { ja: "倍率：", en: "Scale:" },
            scaleUnit: { ja: "%", en: "%" },
            gap: { ja: "間隔：", en: "Gap:" },
            gapUnit: { ja: "pt", en: "pt" },
            columns: { ja: "列数：", en: "Columns:" },
            columnsAuto: { ja: "自動", en: "Auto" },
            errorDetails: { ja: "詳細：", en: "Details:" }
        },
        crop: {
            art: { ja: "アート", en: "Art" },
            trim: { ja: "トリミング", en: "Trim" },
            crop: { ja: "仕上がり", en: "Crop" },
            bleed: { ja: "裁ち落とし", en: "Bleed" }
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
            estimate: { ja: "推定の 行 × 列", en: "Estimated rows × columns" },
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

    /* ラベル（ja/en のリーフ）を現在の言語に解決し、{slash} を / に置換 */
    /* Resolve a label leaf (ja/en) to the current language; {slash} becomes / */
    function L(entry) {
        if (!entry) return "";
        var text = entry[currentLanguage] || entry.en || entry.ja || "";
        return String(text).replace(/\{slash\}/g, "/");
    }

    // =========================================
    // 単位 / Unit
    // =========================================

    // 単位コードとラベルのマップ / Map of unit codes to labels
    var unitLabelMap = {
        0: "in",
        1: "mm",
        2: "pt",
        3: "pica",
        4: "cm",
        5: "Q/H",
        6: "px",
        7: "ft/in",
        8: "m",
        9: "yd",
        10: "ft"
    };

    // 定規単位コードを pt 換算係数へ / Convert a ruler unit code to a pt factor
    function getPtFactorFromUnitCode(code) {
        switch (code) {
            case 0: return 72.0;                        // in
            case 1: return 72.0 / 25.4;                 // mm
            case 2: return 1.0;                         // pt
            case 3: return 12.0;                        // pica
            case 4: return 72.0 / 2.54;                 // cm
            case 5: return 72.0 / 25.4 * 0.25;          // Q or H
            case 6: return 1.0;                         // px
            case 7: return 72.0 * 12.0;                 // ft/in
            case 8: return 72.0 / 25.4 * 1000.0;        // m
            case 9: return 72.0 * 36.0;                 // yd
            case 10: return 72.0 * 12.0;                // ft
            default: return 1.0;
        }
    }

    // 現在の定規単位コードを取得（失敗時は pt）/ Get the current ruler unit code (pt on failure)
    function getCurrentRulerUnitCode() {
        try {
            return app.preferences.getIntegerPreference("rulerType");
        } catch (e) {
            return 2;
        }
    }

    // 現在の定規単位ラベルを取得 / Get the current ruler unit label
    function getCurrentUnitLabel() {
        var unitCode = getCurrentRulerUnitCode();
        return unitLabelMap[unitCode] || "pt";
    }

    // =========================================
    // レイアウト / Layout
    // =========================================

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

    // 例外オブジェクトから表示用の詳細文字列を取り出す / Extract a display detail string from an error
    function SC_getErrorDetailText(e) {
        if (e === undefined || e === null) return "";
        try {
            if (typeof e === "string") return e;
            if (e && e.message) return String(e.message);
            return String(e);
        } catch (e) {
            return "";
        }
    }

    /* ラベルエントリと任意のエラーをアラート表示 / Show an alert from a label entry and an optional error */
    function SC_alert(entry, e) {
        try {
            var msg = L(entry);
            var detail = SC_getErrorDetailText(e);
            if (detail) msg += "\n\n" + L(LABELS.label.errorDetails) + "\n" + detail;
            alert(msg);
        } catch (e) { }
    }

    // ========================
    // 配置処理ヘルパー
    // - 配置前のページ番号指定
    // - 配置サイズの計測
    // - アートボードの作成 / 更新
    // - 単ページの配置
    // - 全アートボードが見える表示倍率への調整
    // ========================

    // PDF 取り込みページ番号を環境設定にセット / Set the PDF import page-number preference
    function placementSetImportPageNumber(pageNum) {
        var n = parseInt(pageNum, 10);
        if (isNaN(n) || n < 1) n = 1;
        try {
            app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", n);
        } catch (e) { }
    }

    // 取り込みページ番号を 1 に戻す / Reset the import page number to 1
    function placementResetImportPageNumber() {
        try {
            app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", 1);
        } catch (e) { }
    }


    // 指定ページを一時配置してサイズを計測 / Temporarily place a page to measure its size
    function placementMeasurePlacedPageSize(doc, fileObj, pageNum, cropMode) {
        if (isPdfLikeFile(fileObj)) {
            SC_setPdfCropPreference(cropMode);
        }
        placementSetImportPageNumber(pageNum);

        var measureItem = null;
        try {
            measureItem = doc.placedItems.add();
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

    // 1 枚目はアクティブABを再利用、以降は新規追加 / Reuse the active artboard first, add new ones afterward
    function placementUseOrAddArtboard(doc, activeIdx, abRect, abCount) {
        if (abCount === 0) {
            doc.artboards[activeIdx].artboardRect = abRect;
        } else {
            doc.artboards.add(abRect);
        }
        return abCount + 1;
    }

    // ページを配置（必要なら倍率を適用）/ Place a page, optionally scaling it
    function placementPlacePage(doc, fileObj, pageNum, pos, cropMode, scalePct) {
        if (isPdfLikeFile(fileObj)) {
            SC_setPdfCropPreference(cropMode);
        }
        placementSetImportPageNumber(pageNum);

        var item = doc.placedItems.add();
        item.file = fileObj;

        if (typeof scalePct === "number" && scalePct !== 100) {
            item.resize(scalePct, scalePct);
        }
        // resize() は中心基準でスケールするため、スケール後に左上位置を確定させる
        // resize() scales about the center, so set the top-left position after scaling
        item.position = pos;
        return item;
    }

    // 1 ページ分のアートボードを用意して配置 / Prepare an artboard and place a single page
    function placementPlaceSinglePage(doc, fileObj, pageNum, pageW, pageH, nextX, baseTop, activeIdx, abCount, cropMode, abGap) {
        var singleRect = [nextX, baseTop, nextX + pageW, baseTop - pageH];
        abCount = placementUseOrAddArtboard(doc, activeIdx, singleRect, abCount);
        var item = placementPlacePage(doc, fileObj, pageNum, [nextX, baseTop], cropMode, 100);
        return {
            nextX: nextX + pageW + abGap,
            abCount: abCount,
            item: item
        };
    }


    // 全アートボードが見えるよう表示を調整 / Fit the view so all artboards are visible
    function placementFitAllArtboardsInView(targetDoc) {
        if (!targetDoc || !targetDoc.artboards || targetDoc.artboards.length === 0) return;

        try {
            app.activeDocument = targetDoc;
        } catch (e) { }

        try {
            var view = targetDoc.activeView;
            if (!view) return;

            var unionLeft = null;
            var unionTop = null;
            var unionRight = null;
            var unionBottom = null;

            for (var i = 0; i < targetDoc.artboards.length; i++) {
                var r = targetDoc.artboards[i].artboardRect;
                if (unionLeft === null || r[0] < unionLeft) unionLeft = r[0];
                if (unionTop === null || r[1] > unionTop) unionTop = r[1];
                if (unionRight === null || r[2] > unionRight) unionRight = r[2];
                if (unionBottom === null || r[3] < unionBottom) unionBottom = r[3];
            }

            if (unionLeft === null) return;

            var unionWidth = unionRight - unionLeft;
            var unionHeight = unionTop - unionBottom;
            if (unionWidth <= 0 || unionHeight <= 0) return;

            var currentBounds = view.bounds;
            var currentWidth = currentBounds[2] - currentBounds[0];
            var currentHeight = currentBounds[1] - currentBounds[3];
            var currentZoom = view.zoom;
            if (!(currentWidth > 0) || !(currentHeight > 0) || !(currentZoom > 0)) return;

            var paddingScale = 0.9;
            var zoomX = currentZoom * (currentWidth / unionWidth);
            var zoomY = currentZoom * (currentHeight / unionHeight);
            var targetZoom = Math.min(zoomX, zoomY) * paddingScale;

            view.centerPoint = [
                (unionLeft + unionRight) / 2,
                (unionTop + unionBottom) / 2
            ];
            view.zoom = targetZoom;
        } catch (e) { }
    }

    // 複数アイテムの可視境界の和を求める / Compute the union of the items' visible bounds
    function placementGetItemsVisibleBounds(items) {
        if (!items || items.length === 0) return null;

        var left = null;
        var top = null;
        var right = null;
        var bottom = null;

        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (!it) continue;

            var b = null;
            try {
                b = it.visibleBounds;
            } catch (e) {
                try {
                    b = it.geometricBounds;
                } catch (e) { }
            }
            if (!b || b.length !== 4) continue;

            if (left === null || b[0] < left) left = b[0];
            if (top === null || b[1] > top) top = b[1];
            if (right === null || b[2] > right) right = b[2];
            if (bottom === null || b[3] < bottom) bottom = b[3];
        }

        if (left === null) return null;
        return [left, top, right, bottom];
    }

    // 指定アイテム全体が見えるよう表示を調整 / Fit the view to the given items
    function placementFitItemsInView(targetDoc, items) {
        if (!targetDoc || !items || items.length === 0) return;

        try {
            app.activeDocument = targetDoc;
        } catch (e) { }

        try {
            var view = targetDoc.activeView;
            if (!view) return;

            var bounds = placementGetItemsVisibleBounds(items);
            if (!bounds) return;

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
            var targetZoom = Math.min(zoomX, zoomY) * paddingScale;

            view.centerPoint = [
                (bounds[0] + bounds[2]) / 2,
                (bounds[1] + bounds[3]) / 2
            ];
            view.zoom = targetZoom;
        } catch (e) { }
    }


    // 最大キャンバス範囲 [left, top, right, bottom] を取得（OMOTI 氏のアイデア）/ Get the largest canvas bounds (idea by OMOTI)
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

    // レイアウト全体の外接サイズ（幅・高さ）を計測。配置と同じ折り返しロジックを使用
    // Measure the overall layout size (width/height) using the same wrapping logic as placement
    function placementMeasureLayoutExtent(targetDoc, fileObj, targetPages, cropMode, scaleFactor, abGap, colsPerRow) {
        var nextX = 0;
        var topY = 0;
        var colCount = 0;
        var rowMaxH = 0;
        var maxRight = 0;
        var minBottom = 0;

        for (var i = 0; i < targetPages.length; i++) {
            var pageNum = parseInt(targetPages[i], 10);
            if (isNaN(pageNum) || pageNum < 1) pageNum = 1;

            var pageW, pageH;
            try {
                var pageSize = placementMeasurePlacedPageSize(targetDoc, fileObj, pageNum, cropMode);
                pageW = pageSize.width * scaleFactor;
                pageH = pageSize.height * scaleFactor;
            } catch (e) {
                continue; // 計測できないページはスキップ / Skip pages that cannot be measured
            }

            var wrapByCol = colsPerRow > 0 && colCount >= colsPerRow;
            var wrapByEdge = colsPerRow === 0 && nextX > 0 && nextX + pageW > MAX_AUTO_ROW_WIDTH;
            if (wrapByCol || wrapByEdge) {
                topY -= rowMaxH + abGap;
                nextX = 0;
                rowMaxH = 0;
                colCount = 0;
            }

            if (nextX + pageW > maxRight) maxRight = nextX + pageW;
            if (topY - pageH < minBottom) minBottom = topY - pageH;

            nextX += pageW + abGap;
            colCount++;
            if (pageH > rowMaxH) rowMaxH = pageH;
        }

        return { width: maxRight, height: -minBottom };
    }

    // =========================================
    // ケイ処理ヘルパー
    // =========================================

    // 配置アイテムの外接矩形でクリッピングマスクグループを作成 / Create a clipping mask group sized to the placed item
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

    // 角丸 LiveEffect を適用 / Apply the Round Corners live effect
    function keiApplyRoundCornersLiveEffect(targetItem, radius) {
        if (!targetItem) return;
        var r = Number(radius);
        if (isNaN(r) || r <= 0) return;
        var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + r + ' "/></LiveEffect>';
        try {
            targetItem.applyEffect(xml);
        } catch (e) { }
    }

    // 配置アイテムにケイ処理を適用 / Apply stroke (kei) processing to a placed item
    // keiOpts: { mode: 'clipGroup', roundCorners: bool, roundRadius: number }
    function keiApplyToPlacedItem(doc, placedItem, keiOpts) {
        if (!keiOpts || !placedItem) return placedItem;

        if (keiOpts.mode === 'clipGroup') {
            var group = keiCreateClippingMaskGroup(placedItem);

            if (keiOpts.roundCorners && keiOpts.roundRadius > 0) {
                keiApplyRoundCornersLiveEffect(group, keiOpts.roundRadius);
            }

            doc.selection = [group];
            app.executeMenuCommand('Adobe New Stroke Shortcut');
            app.executeMenuCommand('Live Pathfinder Exclude');
            doc.selection = null;
            return group;
        }

        return placedItem;
    }

    // =========================================
    // 配置時の仕上がり設定
    // Art / Trim / Bleed / Crop に対応
    // 現在は UI 上では無効だが、内部ロジックは保持
    // =========================================

    // UI用仕上がり設定定数
    var CROP_ART = 4;
    var CROP_TRIM = 3;
    var CROP_BLEED = 2;
    var CROP_CROP = 1;

    // ============================================================
    // ページ数取得ヘルパー
    // - リンクされた PDF/AI の総ページ数（最終ページ番号）を推定
    // - ファイルが明示指定された場合は、一時配置して既存の選択ベースの
    //   ページ数取得ロジックを再利用し、取得後すぐに削除
    // - 現在の選択に対してリンク変更や内容変更は行わない
    // ============================================================

    // 選択中の配置画像のリンク元から総ページ数を取得 / Get the total page count from the selected placed image's link
    function getLastPageFromSelection(selectionItems) {
        var placed = pageCountFindFirstPlacedItem(selectionItems);
        if (!placed) return null;

        var f = placed.file;
        if (!f) {
            SC_alert(LABELS.alert.linkUnknown);
            return null;
        }

        var name = decodeURIComponent(f.name);
        if (!/\.(?:pdf|ai)$/i.test(name)) return null;

        var last = getPageLengthFromFile(f);
        if (!last || isNaN(Number(last)) || Number(last) <= 0) {
            SC_alert(LABELS.alert.pageCountFail);
            return null;
        }

        return Number(last);
    }

    // 指定ファイル（または選択画像）から総ページ数を取得 / Get the total page count from a file or the current selection
    function updatePageCountFromPlacedOrFile(doc, fileObjOrNull, setPathTextFn) {
        var placedTemp = null;

        try {
            // ファイル指定時は一時配置して、既存の選択ベースのページ数取得ロジックを再利用
            if (fileObjOrNull) {
                var name = decodeURIComponent(fileObjOrNull.name);
                if (!/\.(?:pdf|ai)$/i.test(name)) {
                    SC_alert(LABELS.alert.pickPdfAi);
                    return null;
                }

                try {
                    placedTemp = doc.placedItems.add();
                    placedTemp.file = fileObjOrNull;

                    // 一時配置物は、いったん表示中ビューの左上付近へ置く
                    try {
                        var vb = doc.activeView && doc.activeView.bounds ? doc.activeView.bounds : null;
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
            var last = getLastPageFromSelection(doc.selection);

            // 選択中の配置画像があれば、そのファイルパス表示も更新
            try {
                var placedSel = pageCountFindFirstPlacedItem(doc.selection);
                if (placedSel && placedSel.file && setPathTextFn) setPathTextFn(placedSel.file);
            } catch (e) { }

            return last;
        } catch (e) {
            SC_alert(LABELS.alert.pageCountFail, e);
            return null;
        }
    }

    // 選択範囲から最初の PlacedItem を再帰的に探す / Recursively find the first PlacedItem in a selection
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

    // PDF/AI ファイルから総ページ数を推定 / Estimate the total page count from a PDF/AI file
    function getPageLengthFromFile(file) {
        var countInDictPattern = /<<\/Count\s(\d+)/;
        var pageRefPattern1 = /<<\/Type\/Page\/Parent/;
        var pageRefPattern2 = /\/Type\s\/Page\s/;
        var pageRefPattern3 = /\/StructParents\s\d+.*\/Type\/Page>>/;
        var linearizedCountPattern = /<<\/Linearized\s.+\/N\s(\d+)\/T\s.+>>/;
        var pagesNodePattern = /\/Type\/Pages/;
        var countPattern = /\/Count\s(\d+)/;

        var pageCount, line, pageRefCount = 0, declaredCount = 0;
        try {
            file.open('r');
            while (!file.eof) {
                line = file.readln();
                if (countInDictPattern.test(line) || linearizedCountPattern.test(line)) { pageCount = Number(RegExp.$1); break; }
                if (pageRefPattern1.test(line) || pageRefPattern2.test(line) || pageRefPattern3.test(line)) pageRefCount++;
                if (pagesNodePattern.test(line)) {
                    line = file.readln();
                    if (countPattern.test(line)) { declaredCount = Number(RegExp.$1); if (pageRefCount < declaredCount) pageRefCount = declaredCount; }
                }
            }
            if (pageRefCount > 0) pageCount = pageRefCount;
        } catch (e) {
            SC_alert(LABELS.alert.pageCountFail, e);
        } finally {
            try { file.close(); } catch (e) { }
        }
        return pageCount;
    }

    // PDF 配置時の crop プリファレンスキー（環境差を吸収するため複数試行）/ Crop preference keys (multiple keys tried for version differences)
    var PDF_CROP_PREFERENCE_KEYS = [
        "plugin/PDFImport/CropToBox",
        "plugin/PDFImport/CropTo",
        "plugin/PDFImport/CropBox",
        "plugin/PDFImport/CropToType"
    ];

    // crop 種別を複数キーに設定（環境差を吸収）/ Set the crop type across several keys (absorbs version differences)
    function SC_setPdfCropPreference(cropVal) {
        for (var i = 0; i < PDF_CROP_PREFERENCE_KEYS.length; i++) {
            try {
                app.preferences.setIntegerPreference(PDF_CROP_PREFERENCE_KEYS[i], cropVal);
            } catch (e) { }
        }
    }

    // 現在の crop プリファレンスを退避 / Snapshot the current crop preferences
    function SC_snapshotPdfCropPreference() {
        var snapshot = {};
        for (var i = 0; i < PDF_CROP_PREFERENCE_KEYS.length; i++) {
            try {
                snapshot[PDF_CROP_PREFERENCE_KEYS[i]] = app.preferences.getIntegerPreference(PDF_CROP_PREFERENCE_KEYS[i]);
            } catch (e) { }
        }
        return snapshot;
    }

    // 退避した crop プリファレンスを復元 / Restore the snapshotted crop preferences
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

    /**
     * Illustrator のバージョン差異を吸収するため、複数キーに試行します。
     * 期待する値（多くの環境で）: 0=Media, 1=Crop, 2=Bleed, 3=Trim, 4=Art
     */

    // 拡張子から PDF/AI ファイルかどうかを判定 / Heuristically detect a PDF/AI file by extension
    function isPdfLikeFile(f) {
        var n = String((f && f.name) || "").toLowerCase();
        return (n.indexOf(".pdf") > -1) || (n.indexOf(".ai") > -1);
    }


    // 入力文字列（例: "1-20", "1,3,5"）をページ番号配列へ変換し、重複と範囲外（1..maxPages 以外）を除外
    // Parse a page-range string into page numbers, dropping duplicates and out-of-range values (outside 1..maxPages)
    function parsePageNumbers(inputStr, maxPages) {
        var result = [];
        var seen = {};
        var hasMax = (typeof maxPages === "number" && maxPages > 0);

        // 1 件のページ番号を検証して追加 / Validate and append a single page number
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

    function main() {

        if (app.documents.length === 0) {
            SC_alert(LABELS.alert.needDoc);
            return;
        }

        var srcDoc = app.activeDocument;
        var doc = srcDoc;
        // 現在の読み込み対象ファイル。未選択時は null。
        var sourceFile = null;
        var detectedRangeText = "";
        var customRangeText = "";

        // ------------------------
        // UI構築
        // ------------------------
        function buildDialogUI() {
            var win = new Window("dialog", L(LABELS.dialog.title) + " " + SCRIPT_VERSION);
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

            var sourcePanel = leftColumnGroup.add('panel', undefined, L(LABELS.panel.source));
            setupPanel(sourcePanel);
            var btnBrowse = sourcePanel.add('button', undefined, L(LABELS.button.selectFile));
            btnBrowse.helpTip = L(LABELS.tooltip.source);
            var stSourceName = sourcePanel.add('statictext', undefined, L(LABELS.label.notSelected));
            stSourceName.characters = 16;

            var pagesPanel = leftColumnGroup.add('panel', undefined, L(LABELS.panel.pages));
            setupPanel(pagesPanel);
            var rangeModeGroup = pagesPanel.add('group');
            setupGroup(rangeModeGroup, 'column');
            var rbRangeAll = rangeModeGroup.add('radiobutton', undefined, L(LABELS.radio.rangeAll));
            var rbRangeFirst = rangeModeGroup.add('radiobutton', undefined, L(LABELS.radio.rangeFirst));
            var rbRangeCustom = rangeModeGroup.add('radiobutton', undefined, L(LABELS.radio.rangeCustom));
            // 初期状態は「全ページ」を選択
            rbRangeAll.value = true;

            var rangeInputGroup = pagesPanel.add('group');
            setupGroup(rangeInputGroup, 'row');
            var etRange = rangeInputGroup.add('edittext', undefined, '');
            etRange.characters = 10;
            etRange.enabled = true;
            etRange.helpTip = L(LABELS.tooltip.customRange);

            var totalPagesGroup = pagesPanel.add('group');
            totalPagesGroup.orientation = 'row';
            totalPagesGroup.alignChildren = ['right', 'center'];
            totalPagesGroup.alignment = ['fill', 'top'];
            totalPagesGroup.add('statictext', undefined, '');
            var stTotalPages = totalPagesGroup.add('statictext', undefined, '');
            stTotalPages.characters = 8;
            stTotalPages.justify = 'right';
            stTotalPages.helpTip = L(LABELS.tooltip.totalPages);

            var methodPanel = leftColumnGroup.add("panel", undefined, L(LABELS.panel.mode));
            setupPanel(methodPanel);
            var methodGroup = methodPanel.add("group");
            setupGroup(methodGroup, "column");
            var rbPerArtboard = methodGroup.add("radiobutton", undefined, L(LABELS.radio.perArtboard));
            rbPerArtboard.helpTip = L(LABELS.tooltip.perArtboard);
            var rbIgnoreArtboard = methodGroup.add("radiobutton", undefined, L(LABELS.radio.ignoreArtboard));
            rbIgnoreArtboard.helpTip = L(LABELS.tooltip.ignoreArtboard);
            rbPerArtboard.value = true;

            var layoutPanel = rightColumnGroup.add("panel", undefined, L(LABELS.panel.placement));
            setupPanel(layoutPanel);

            var columnsGroup = layoutPanel.add("group");
            setupGroup(columnsGroup, "row");
            var stColsLabel = columnsGroup.add("statictext", undefined, L(LABELS.label.columns));
            var etCols = columnsGroup.add("edittext", undefined, L(LABELS.label.columnsAuto));
            etCols.characters = 5;
            etCols.helpTip = L(LABELS.tooltip.columns);
            stColsLabel.helpTip = L(LABELS.tooltip.columns);
            // 推定レイアウト（行 × 列）は列数フィールドの右にインライン表示 / Show the estimated rows × columns inline, right of the column field
            var stEstimateLayout = columnsGroup.add("statictext", undefined, "");
            stEstimateLayout.characters = 10;
            stEstimateLayout.helpTip = L(LABELS.tooltip.estimate);

            var gapGroup = layoutPanel.add("group");
            setupGroup(gapGroup, "row");
            var stGapLabel = gapGroup.add("statictext", undefined, L(LABELS.label.gap));
            var etArtboardGap = gapGroup.add("edittext", undefined, String(DEFAULT_ARTBOARD_GAP));
            etArtboardGap.characters = 5;
            var stGapUnit = gapGroup.add("statictext", undefined, L(LABELS.label.gapUnit));

            var scaleGroup = layoutPanel.add("group");
            setupGroup(scaleGroup, "row");
            var stScaleLabel = scaleGroup.add("statictext", undefined, L(LABELS.label.scale));
            var etScale = scaleGroup.add("edittext", undefined, "100");
            etScale.characters = 5;
            etScale.helpTip = L(LABELS.tooltip.scale);
            stScaleLabel.helpTip = L(LABELS.tooltip.scale);
            var stScaleUnit = scaleGroup.add("statictext", undefined, L(LABELS.label.scaleUnit));
            var ddCrop = layoutPanel.add("dropdownlist", undefined, [L(LABELS.crop.art), L(LABELS.crop.trim), L(LABELS.crop.crop), L(LABELS.crop.bleed)]);
            ddCrop.minimumSize.width = 160;
            ddCrop.selection = 2;
            ddCrop.enabled = false;

            var keiPanel = rightColumnGroup.add("panel", undefined, L(LABELS.panel.kei));
            setupPanel(keiPanel);
            var keiModeGroup = keiPanel.add("group");
            setupGroup(keiModeGroup, "column");
            var rbKeiNone = keiModeGroup.add("radiobutton", undefined, L(LABELS.radio.keiNone));
            var rbKeiClipGroup = keiModeGroup.add("radiobutton", undefined, L(LABELS.radio.keiClipGroup));
            rbKeiNone.value = true;

            var roundCornerGroup = keiPanel.add("group");
            setupGroup(roundCornerGroup, "row");
            var cbRoundCorner = roundCornerGroup.add("checkbox", undefined, L(LABELS.checkbox.roundCorner));
            var etRoundCorner = roundCornerGroup.add("edittext", undefined, "3");
            etRoundCorner.characters = 5;
            var stRoundCornerUnit = roundCornerGroup.add("statictext", undefined, getCurrentUnitLabel());
            cbRoundCorner.value = false;
            etRoundCorner.enabled = false;
            cbRoundCorner.helpTip = L(LABELS.tooltip.roundCorner);
            etRoundCorner.helpTip = L(LABELS.tooltip.roundCorner);

            var progressBar = win.add("progressbar", undefined, 0, 100);
            progressBar.preferredSize.width = 300;
            progressBar.preferredSize.height = 8;
            progressBar.visible = false;

            var buttonGroup = win.add("group");
            buttonGroup.orientation = "row";
            buttonGroup.alignChildren = ["center", "center"];
            var btnCancel = buttonGroup.add("button", undefined, L(LABELS.button.cancel), { name: "cancel" });
            var btnOk = buttonGroup.add("button", undefined, L(LABELS.button.ok), { name: "ok" });

            return {
                win: win,
                btnBrowse: btnBrowse,
                stSourceName: stSourceName,
                stTotalPages: stTotalPages,
                pagesPanel: pagesPanel,
                methodPanel: methodPanel,
                layoutPanel: layoutPanel,
                keiPanel: keiPanel,
                rbRangeAll: rbRangeAll,
                rbRangeFirst: rbRangeFirst,
                rbRangeCustom: rbRangeCustom,
                etRange: etRange,
                rbPerArtboard: rbPerArtboard,
                rbIgnoreArtboard: rbIgnoreArtboard,
                stScaleLabel: stScaleLabel,
                stScaleUnit: stScaleUnit,
                stGapLabel: stGapLabel,
                stGapUnit: stGapUnit,
                etCols: etCols,
                etScale: etScale,
                etArtboardGap: etArtboardGap,
                ddCrop: ddCrop,
                stEstimateLayout: stEstimateLayout,
                rbKeiNone: rbKeiNone,
                rbKeiClipGroup: rbKeiClipGroup,
                cbRoundCorner: cbRoundCorner,
                etRoundCorner: etRoundCorner,
                stRoundCornerUnit: stRoundCornerUnit,
                progressBar: progressBar,
                btnCancel: btnCancel,
                btnOk: btnOk
            };
        }

        // ダイアログUIを構築し、主要な参照を取り出す
        var ui = buildDialogUI();
        var win = ui.win;
        var btnBrowse = ui.btnBrowse;
        var stSourceName = ui.stSourceName;
        var stTotalPages = ui.stTotalPages;
        var pagesPanel = ui.pagesPanel;
        var methodPanel = ui.methodPanel;
        var layoutPanel = ui.layoutPanel;
        var keiPanel = ui.keiPanel;
        var rbRangeAll = ui.rbRangeAll;
        var rbRangeFirst = ui.rbRangeFirst;
        var rbRangeCustom = ui.rbRangeCustom;
        var etRange = ui.etRange;
        var rbPerArtboard = ui.rbPerArtboard;
        var rbIgnoreArtboard = ui.rbIgnoreArtboard;
        var stScaleLabel = ui.stScaleLabel;
        var stScaleUnit = ui.stScaleUnit;
        var stGapLabel = ui.stGapLabel;
        var stGapUnit = ui.stGapUnit;
        var etCols = ui.etCols;
        var etScale = ui.etScale;
        var etArtboardGap = ui.etArtboardGap;
        var ddCrop = ui.ddCrop;
        var stEstimateLayout = ui.stEstimateLayout;
        var rbKeiNone = ui.rbKeiNone;
        var rbKeiClipGroup = ui.rbKeiClipGroup;
        var cbRoundCorner = ui.cbRoundCorner;
        var etRoundCorner = ui.etRoundCorner;
        var stRoundCornerUnit = ui.stRoundCornerUnit;
        var progressBar = ui.progressBar;
        var btnCancel = ui.btnCancel;
        var btnOk = ui.btnOk;

        function updatePanelEnabledState() {
            var enabled = !!sourceFile;
            pagesPanel.enabled = enabled;
            methodPanel.enabled = enabled;
            layoutPanel.enabled = enabled;
            keiPanel.enabled = enabled;
        }

        function updateRangeEnabledState() {
            if (rbRangeCustom.value) {
                etRange.enabled = true;
                etRange.text = customRangeText || detectedRangeText || '';
            } else {
                if (etRange.enabled) {
                    customRangeText = etRange.text;
                }
                etRange.enabled = false;
                etRange.text = '';
            }
        }

        function updateTotalPagesLabel() {
            var totalPages = 0;
            var placedPages = 0;

            totalPages = getDetectedTotalPages();

            if (rbRangeAll.value) {
                placedPages = totalPages;
            } else if (rbRangeFirst.value) {
                placedPages = totalPages > 0 ? 1 : 0;
            } else {
                placedPages = parsePageNumbers(customRangeText || detectedRangeText || '', totalPages).length;
            }

            stTotalPages.text = totalPages > 0 ? (placedPages + ' / ' + totalPages) : '';
        }

        function getDetectedTotalPages() {
            if (!detectedRangeText) return 0;
            var m = detectedRangeText.match(/^(?:1-)?(\d+)$/);
            return m ? (parseInt(m[1], 10) || 0) : 0;
        }

        function updatePlacementEstimateInfo() {
            var totalPages = 0;
            var placedPages = 0;
            var colsPerRow = getColsPerRow();

            totalPages = getDetectedTotalPages();

            if (rbRangeAll.value) {
                placedPages = totalPages;
            } else if (rbRangeFirst.value) {
                placedPages = totalPages > 0 ? 1 : 0;
            } else {
                placedPages = parsePageNumbers(customRangeText || detectedRangeText || '', totalPages).length;
            }

            if (colsPerRow > 0) {
                var actualCols = placedPages > 0 ? Math.min(colsPerRow, placedPages) : 0;
                var actualRows = placedPages > 0 ? Math.ceil(placedPages / colsPerRow) : 0;
                stEstimateLayout.text = actualRows + ' × ' + actualCols;
            } else {
                stEstimateLayout.text = '';
            }
        }

        // ------------------------
        // UI表示補助
        // ------------------------
        function getFileDisplayInfo(f) {
            if (!f) {
                return {
                    name: L(LABELS.label.notSelected),
                    path: ''
                };
            }

            var nameText;
            var pathText;

            try {
                nameText = decodeURIComponent(f.name);
            } catch (e) {
                nameText = String(f.name || L(LABELS.label.notSelected));
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

        function isPdfOrAiFile(fileObj) {
            if (!fileObj) return false;
            if (fileObj instanceof Folder) return true;
            var name = String(fileObj.name || "").toLowerCase();
            return /\.(pdf|ai)$/i.test(name);
        }

        function setPathText(f) {
            // 現在の読み込み対象ファイル参照を更新
            sourceFile = f || null;
            updatePanelEnabledState();

            var info = getFileDisplayInfo(f);
            stSourceName.text = info.name;
            stSourceName.helpTip = info.path;
        }

        function changeValueByArrowKey(editText) {
            editText.addEventListener("keydown", function (event) {
                var textValue = String(editText.text);
                var value;

                if (textValue === L(LABELS.label.columnsAuto)) {
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
        // UIイベント配線
        // ------------------------
        function bindDialogEvents() {
            btnBrowse.onClick = function () {
                var f = File.openDialog(L(LABELS.dialog.pickFile), isPdfOrAiFile);
                var last;
                if (!f) return;
                last = updatePageCountFromPlacedOrFile(srcDoc, f, setPathText);
                detectedRangeText = last ? ('1-' + last) : '';
                updateTotalPagesLabel();
                if (!customRangeText) {
                    customRangeText = detectedRangeText;
                }
                updateRangeEnabledState();
                updatePlacementEstimateInfo();
            };

            rbRangeAll.onClick = function () {
                updateRangeEnabledState();
                updateTotalPagesLabel();
                updatePlacementEstimateInfo();
            };

            rbRangeFirst.onClick = function () {
                updateRangeEnabledState();
                updateTotalPagesLabel();
                updatePlacementEstimateInfo();
            };

            rbRangeCustom.onClick = function () {
                updateRangeEnabledState();
                updateTotalPagesLabel();
                updatePlacementEstimateInfo();
            };

            etRange.onChanging = function () {
                if (rbRangeCustom.value) {
                    customRangeText = etRange.text;
                    updateTotalPagesLabel();
                    updatePlacementEstimateInfo();
                }
            };

            etRange.onChange = function () {
                if (rbRangeCustom.value) {
                    customRangeText = etRange.text;
                    updateTotalPagesLabel();
                    updatePlacementEstimateInfo();
                }
            };

            etCols.onChanging = function () {
                updatePlacementEstimateInfo();
            };

            etCols.onChange = function () {
                updatePlacementEstimateInfo();
            };

            rbPerArtboard.onClick = function () {
                updateScaleEnabledState();
                updateGapHelpTip();
                updatePlacementEstimateInfo();
            };

            rbIgnoreArtboard.onClick = function () {
                updateScaleEnabledState();
                updateGapHelpTip();
                updatePlacementEstimateInfo();
            };

            function updateKeiRoundEnabled() {
                var keiActive = !rbKeiNone.value;
                cbRoundCorner.enabled = keiActive;
                etRoundCorner.enabled = keiActive && cbRoundCorner.value;
            }

            rbKeiNone.onClick = function () { updateKeiRoundEnabled(); };
            rbKeiClipGroup.onClick = function () { updateKeiRoundEnabled(); };
            cbRoundCorner.onClick = function () { updateKeiRoundEnabled(); };
            updateKeiRoundEnabled();

            btnCancel.onClick = function () {
                win.close(2);
            };

            btnOk.onClick = function () {
                if (!sourceFile) {
                    SC_alert(LABELS.alert.needFile);
                    return;
                }
                var finalPages;
                if (rbRangeAll.value) {
                    var lastPage = getDetectedTotalPages();
                    if (!lastPage || lastPage < 1) lastPage = 1;
                    finalPages = [];
                    for (var p = 1; p <= lastPage; p++) {
                        finalPages.push(p);
                    }
                } else if (rbRangeFirst.value) {
                    finalPages = [1];
                } else {
                    finalPages = parsePageNumbers(etRange.text, getDetectedTotalPages());
                    if (!finalPages || finalPages.length === 0) finalPages = [1];
                }

                var cropMode = getCropModeFromUI();
                var keiOpts = getKeiOptionsFromUI();
                if (rbPerArtboard.value) {
                    placeOnIndividualArtboards(finalPages, cropMode, keiOpts);
                } else {
                    var scale = getScaleFromUI();
                    placeIgnoringArtboards(finalPages, cropMode, scale, keiOpts);
                }
                win.close(1);
            };
        }

        function updateScaleEnabledState() {
            var enabled = !rbPerArtboard.value;
            etScale.enabled = enabled;
            stScaleLabel.enabled = enabled;
            stScaleUnit.enabled = enabled;
        }

        function updateGapHelpTip() {
            var helpText = rbPerArtboard.value ? L(LABELS.tooltip.gapPerArtboard) : L(LABELS.tooltip.gapIgnoreArtboard);
            stGapLabel.helpTip = helpText;
            etArtboardGap.helpTip = helpText;
            stGapUnit.helpTip = helpText;
        }

        // ------------------------
        // 初期状態反映
        // ------------------------
        function initializeDialogState() {
            var last;
            ddCrop.enabled = false;
            updatePanelEnabledState();
            last = updatePageCountFromPlacedOrFile(srcDoc, null, setPathText);
            detectedRangeText = last ? ('1-' + last) : '';
            updateTotalPagesLabel();
            if (!customRangeText) {
                customRangeText = detectedRangeText;
            }
            updateRangeEnabledState();
            updateScaleEnabledState();
            updateGapHelpTip();
            updatePlacementEstimateInfo();

            stRoundCornerUnit.text = getCurrentUnitLabel();
            changeValueByArrowKey(etCols);
            changeValueByArrowKey(etArtboardGap);
            changeValueByArrowKey(etScale);
        }

        // UIで指定した1行あたりの列数を取得（0 = 自動）
        function getColsPerRow() {
            var v = parseInt(etCols.text, 10);
            if (isNaN(v) || v <= 0) return 0;
            return v;
        }

        // UIで指定したアートボード間隔（pt）を取得
        function getArtboardGap() {
            var v = parseFloat(etArtboardGap.text);
            if (isNaN(v) || v < 0) v = DEFAULT_ARTBOARD_GAP;
            return v;
        }

        // ------------------------
        // 配置オプション取得
        // ------------------------
        function getKeiOptionsFromUI() {
            if (rbKeiNone.value) {
                return null;
            }
            return {
                mode: 'clipGroup',
                roundCorners: cbRoundCorner.value,
                roundRadius: (function () {
                    var n = parseFloat(etRoundCorner.text);
                    if (isNaN(n) || n < 0) n = 0;
                    return n * getPtFactorFromUnitCode(getCurrentRulerUnitCode());
                })()
            };
        }

        function getScaleFromUI() {
            var v = parseFloat(etScale.text);
            if (isNaN(v) || v <= 0) v = 100;
            return v;
        }

        function getCropModeFromUI() {
            var idx = (ddCrop.selection) ? ddCrop.selection.index : 2;
            // 0:アート / 1:トリミング / 2:仕上がり / 3:裁ち落とし
            // 現在は UI 上で変更不可のため、既定の selection をそのまま使う
            if (idx === 0) return CROP_ART;
            if (idx === 1) return CROP_TRIM;
            if (idx === 3) return CROP_BLEED;
            // 仕上がりは CropBox を想定
            return CROP_CROP;
        }

        // ------------------------
        // プログレスバー更新
        // ------------------------
        function showProgress() {
            progressBar.value = 0;
            progressBar.visible = true;
            win.update();
        }

        function updateProgress(current, total) {
            progressBar.value = Math.round((current / total) * 100);
            win.update();
        }

        function hideProgress() {
            progressBar.visible = false;
            win.update();
        }

        // ------------------------
        // 実行処理
        // ------------------------
        function placeOnIndividualArtboards(targetPages, cropMode, keiOpts) {
            var abGap = getArtboardGap();
            var colsPerRow = getColsPerRow();

            // レイアウト全体のサイズを先に計測し、カンバス中心に揃える
            // Measure the whole layout first, then center it on the canvas
            var extent = placementMeasureLayoutExtent(doc, sourceFile, targetPages, cropMode, 1, abGap, colsPerRow);
            var canvas = getLargestCanvasBounds(doc);
            var startX = (canvas[0] + canvas[2]) / 2 - extent.width / 2;
            var baseTop = (canvas[1] + canvas[3]) / 2 + extent.height / 2;

            var activeIdx = doc.artboards.getActiveArtboardIndex();
            var nextX = startX;
            var abCount = 0;
            var colCount = 0;
            var rowMaxH = 0; // 現在の行で最も高いページの高さ / Tallest page height in the current row
            var skippedCount = 0;
            var completed = false;
            var cropSnapshot = SC_snapshotPdfCropPreference();
            showProgress();

            try {
                for (var i = 0; i < targetPages.length; i++) {
                    try {
                        var pageNum = parseInt(targetPages[i], 10);
                        if (isNaN(pageNum) || pageNum < 1) pageNum = 1;

                        var pageSize = placementMeasurePlacedPageSize(doc, sourceFile, pageNum, cropMode);
                        var pageW = pageSize.width;
                        var pageH = pageSize.height;

                        // 指定列数に達した、または行幅が上限を超えるとき次の行へ折り返す
                        // Wrap to the next row when the column count or the row width limit is reached
                        var wrapByCol = colsPerRow > 0 && colCount >= colsPerRow;
                        var wrapByEdge = colsPerRow === 0 && (nextX - startX) > 0 && (nextX - startX) + pageW > MAX_AUTO_ROW_WIDTH;
                        if (wrapByCol || wrapByEdge) {
                            baseTop -= rowMaxH + abGap;
                            nextX = startX;
                            rowMaxH = 0;
                            colCount = 0;
                        }

                        var result = placementPlaceSinglePage(doc, sourceFile, pageNum, pageW, pageH, nextX, baseTop, activeIdx, abCount, cropMode, abGap);

                        if (keiOpts && result.item) {
                            try { keiApplyToPlacedItem(doc, result.item, keiOpts); } catch (e) { }
                        }

                        nextX = result.nextX;
                        abCount = result.abCount;
                        colCount++;
                        if (pageH > rowMaxH) rowMaxH = pageH;
                    } catch (e) {
                        skippedCount++; // 失敗ページはスキップして継続 / Skip the failed page and continue
                    }
                    updateProgress(i + 1, targetPages.length);
                }
                completed = true;
            } catch (e) {
                SC_alert(LABELS.alert.placeError, e);
            } finally {
                try { placementResetImportPageNumber(); } catch (e) { }
                try { SC_restorePdfCropPreference(cropSnapshot); } catch (e) { }
                hideProgress();
            }

            if (completed) {
                placementFitAllArtboardsInView(doc);
                if (skippedCount > 0) SC_alert(LABELS.alert.someSkipped);
            }
        }

        function placeIgnoringArtboards(targetPages, cropMode, scale, keiOpts) {
            var scaleFactor = scale / 100;
            var abGap = getArtboardGap();
            var colsPerRow = getColsPerRow();

            // レイアウト全体のサイズを先に計測し、カンバス中心に揃える
            // Measure the whole layout first, then center it on the canvas
            var extent = placementMeasureLayoutExtent(doc, sourceFile, targetPages, cropMode, scaleFactor, abGap, colsPerRow);
            var canvas = getLargestCanvasBounds(doc);
            var startX = (canvas[0] + canvas[2]) / 2 - extent.width / 2;
            var baseTop = (canvas[1] + canvas[3]) / 2 + extent.height / 2;

            var nextX = startX;
            var placedItems = [];
            var colCount = 0;
            var rowMaxH = 0; // 現在の行で最も高いページの高さ（スケール適用後）/ Tallest page height in the current row (scaled)
            var skippedCount = 0;
            var completed = false;
            var cropSnapshot = SC_snapshotPdfCropPreference();
            showProgress();

            try {
                for (var i = 0; i < targetPages.length; i++) {
                    try {
                        var pageNum = parseInt(targetPages[i], 10);
                        if (isNaN(pageNum) || pageNum < 1) pageNum = 1;

                        var pageSize = placementMeasurePlacedPageSize(doc, sourceFile, pageNum, cropMode);
                        var pageW = pageSize.width * scaleFactor;
                        var pageH = pageSize.height * scaleFactor;

                        // 指定列数に達した、または行幅が上限を超えるとき次の行へ折り返す
                        // Wrap to the next row when the column count or the row width limit is reached
                        var wrapByCol = colsPerRow > 0 && colCount >= colsPerRow;
                        var wrapByEdge = colsPerRow === 0 && (nextX - startX) > 0 && (nextX - startX) + pageW > MAX_AUTO_ROW_WIDTH;
                        if (wrapByCol || wrapByEdge) {
                            baseTop -= rowMaxH + abGap;
                            nextX = startX;
                            rowMaxH = 0;
                            colCount = 0;
                        }

                        var placedItem = placementPlacePage(doc, sourceFile, pageNum, [nextX, baseTop], cropMode, scale);

                        if (keiOpts) {
                            try { placedItem = keiApplyToPlacedItem(doc, placedItem, keiOpts); } catch (e) { }
                        }

                        placedItems.push(placedItem);

                        nextX += pageW + abGap;
                        colCount++;
                        if (pageH > rowMaxH) rowMaxH = pageH;
                    } catch (e) {
                        skippedCount++; // 失敗ページはスキップして継続 / Skip the failed page and continue
                    }
                    updateProgress(i + 1, targetPages.length);
                }
                completed = true;
            } catch (e) {
                SC_alert(LABELS.alert.placeError, e);
            } finally {
                try { placementResetImportPageNumber(); } catch (e) { }
                try { SC_restorePdfCropPreference(cropSnapshot); } catch (e) { }
                hideProgress();
            }

            if (completed) {
                try {
                    doc.selection = null;
                } catch (e) { }
                for (var j = 0; j < placedItems.length; j++) {
                    try {
                        placedItems[j].selected = true;
                    } catch (e) { }
                }
                placementFitItemsInView(doc, placedItems);
                if (skippedCount > 0) SC_alert(LABELS.alert.someSkipped);
            }
        }


        bindDialogEvents();
        initializeDialogState();

        win.show();

    }

    main();

}());