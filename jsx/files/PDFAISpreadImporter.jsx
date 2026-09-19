#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

PDF/AI ファイルを指定したページ範囲で読み込み、新規ドキュメント上に各ページを個別のアートボードとして配置します。
横長ページは見開きとして自動判定し、左右2つのアートボードに分割します。

詳細は README を参照してください。

### Overview

Imports a PDF/AI file over a given page range and places each page on its own artboard in a new document.
Landscape pages are detected as spreads and split into two artboards, left and right.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PDFAISpreadImporter";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PDFAISpreadImporter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PDFAISpreadImporter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n5514d9f2c5f8"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 生成されるアートボード間の間隔（pt） / Gap between generated artboards (pt) */
    var ARTBOARD_GAP = 100;

    /* 新規出力ドキュメントのラスタライズ効果解像度（ppi） / Raster effects resolution of the new document (ppi) */
    var RASTER_EFFECTS_RESOLUTION = 300;

    /* 見開きと判定する横長比（幅 > 高さ × この値） / Aspect ratio that marks a page as a spread */
    var SPREAD_ASPECT_RATIO = 1.2;

    /* 綴じ方向の判定で読むPDFの行数 / Number of PDF lines scanned for the binding direction */
    var BINDING_SCAN_LINE_LIMIT = 200;

    /* 全アートボードを表示するときの余白率 / Padding ratio when fitting all artboards in the view */
    var VIEW_FIT_PADDING_RATIO = 0.9;

    /* PDF読み込み時のトリミング指定値 / Crop-to values used when importing a PDF */
    var CROP_CROP = 1;
    var CROP_BLEED = 2;
    var CROP_TRIM = 3;
    var CROP_ART = 4;

    /* トリミング指定の初期選択（仕上がり） / Default crop option */
    var DEFAULT_CROP_INDEX = 2;

    // =========================================
    // レイアウト設定 / Layout settings
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10];
    var BUTTON_ROW_TOP_MARGIN = 10;
    var SOURCE_NAME_CHARS = 16;
    var PAGE_RANGE_CHARS = 10;
    var CROP_LIST_MIN_WIDTH = 160;

    // =========================================
    // 多言語ラベル / Localized labels
    // =========================================

    /**
     * UIの表示言語を判定します。
     * @returns {string} "ja" または "en"
     */
    function getUiLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getUiLang();

    var LABELS = {
        dialog: {
            title: { ja: "PDF/AI見開き配置", en: "PDF/AI Spread Placement" },
            pickFile: { ja: "PDF/AIを選択してください", en: "Select a PDF/AI" },
            pickFilter: { ja: "PDF/AI:*.pdf;*.ai", en: "PDF/AI:*.pdf;*.ai" }
        },

        panel: {
            source: { ja: "読み込みファイル", en: "Source File" },
            artboard: { ja: "アートボード", en: "Artboards" },
            placement: { ja: "配置方法", en: "Placement" },
            newDocument: { ja: "新規ドキュメント", en: "New Document" }
        },

        fieldLabel: {
            pageRange: { ja: "指定", en: "Range" },
            evenPage: { ja: "偶数ページ", en: "Even Pages" },
            colorMode: { ja: "カラーモード", en: "Color Mode" },
            notSelected: { ja: "未指定", en: "Not selected" }
        },

        radio: {
            evenPageRight: { ja: "右", en: "Right" },
            evenPageLeft: { ja: "左", en: "Left" },
            colorModeCMYK: { ja: "CMYK", en: "CMYK" },
            colorModeRGB: { ja: "RGB", en: "RGB" }
        },

        dropdown: {
            cropArt: { ja: "アート", en: "Art" },
            cropTrim: { ja: "トリミング", en: "Trim" },
            cropCrop: { ja: "仕上がり", en: "Crop" },
            cropBleed: { ja: "裁ち落とし", en: "Bleed" }
        },

        button: {
            selectFile: { ja: "ファイル指定", en: "Select File" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },

        tooltip: {
            selectFile: {
                ja: "読み込むPDF/AIファイルを選びます。選択中の配置画像があれば、起動時にそのリンク先を読み込みます。",
                en: "Choose the PDF/AI file to import. A selected placed image is picked up on startup."
            },
            pageRange: {
                ja: "読み込むページを「1-10」「1,3,5」のように指定します。空欄のときは1ページ目だけを配置します。",
                en: "Pages to import, written as 1-10 or 1,3,5. Leave it empty to place only the first page."
            },
            cropMode: {
                ja: "PDFのどの領域を基準に配置するかを選びます。AIファイルでは使用しません。",
                en: "Which PDF box the pages are placed from. Not used for AI files."
            },
            evenPage: {
                ja: "見開きを左右に分割したとき、偶数ページを置く側です。PDFの綴じ方向から自動で設定します。",
                en: "Which side the even pages go to when a spread is split. Detected from the PDF binding direction."
            }
        },

        alert: {
            needDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document before running." },
            needFile: { ja: "先に［ファイル指定］で読み込みファイルを選択してください。", en: "Please select a source file first." },
            pickPdfAi: { ja: "PDFまたはAIファイルを選択してください。", en: "Please select a PDF or AI file." },
            linkUnknown: { ja: "画像のリンク先が不明でした。", en: "Image link not found." },
            pageCountFailed: { ja: "リンクされたPDF/AIファイルのページ数を取得できませんでした。", en: "Could not determine the page count of the linked PDF/AI file." },
            placeFailed: { ja: "配置中にエラーが発生しました。", en: "An error occurred while placing the pages." },
            errorDetails: { ja: "詳細：", en: "Details:" }
        }
    };

    /**
     * ラベル定義から表示言語の文字列を取り出します。
     * @param {object} labelSet - { ja: string, en: string } 形式のラベル定義
     * @returns {string} 表示言語の文字列
     */
    function getLabel(labelSet) {
        if (!labelSet) return "";
        return labelSet[uiLang] || labelSet.en || labelSet.ja || "";
    }

    /* トリミング指定の選択肢。ドロップダウンの並び順と一致させます。 */
    var CROP_OPTIONS = [
        { value: CROP_ART, labelSet: LABELS.dropdown.cropArt },
        { value: CROP_TRIM, labelSet: LABELS.dropdown.cropTrim },
        { value: CROP_CROP, labelSet: LABELS.dropdown.cropCrop },
        { value: CROP_BLEED, labelSet: LABELS.dropdown.cropBleed }
    ];

    // =========================================
    // アラート表示 / Alerts
    // =========================================

    /**
     * 例外オブジェクトから表示用の詳細文字列を取り出します。
     * @param {*} e - 例外オブジェクトまたは文字列
     * @returns {string} 詳細文字列。詳細がなければ空文字
     */
    function getErrorDetailText(e) {
        if (e === undefined || e === null) return "";
        if (typeof e === "string") return e;
        return String((e && e.message) ? e.message : e);
    }

    /**
     * アラートを表示します。例外を渡すと詳細を併記します。
     * @param {object} labelSet - 表示するメッセージのラベル定義
     * @param {*} [e] - 例外オブジェクト
     * @returns {void}
     */
    function showAlert(labelSet, e) {
        var message = getLabel(labelSet);
        var detail = getErrorDetailText(e);
        if (detail) message += "\n\n" + getLabel(LABELS.alert.errorDetails) + "\n" + detail;
        alert(message);
    }

    // =========================================
    // ファイル判定 / File checks
    // =========================================

    /**
     * URLエンコードされた文字列を復号します。復号できない場合は元の文字列を返します。
     * @param {string} text - 復号する文字列
     * @returns {string} 復号後の文字列
     */
    function decodeSafely(text) {
        var source = String(text || "");
        try {
            return decodeURIComponent(source);
        } catch (e) {
            return source;
        }
    }

    /**
     * PDFまたはAIファイルかどうかを拡張子で判定します。
     * @param {File} file - 判定するファイル
     * @returns {boolean} PDFまたはAIなら true
     */
    function isPdfOrAiFile(file) {
        return !!file && /\.(?:pdf|ai)$/i.test(decodeSafely(file.name));
    }

    /**
     * PDFファイルかどうかを拡張子で判定します。
     * @param {File} file - 判定するファイル
     * @returns {boolean} PDFなら true
     */
    function isPdfFile(file) {
        return !!file && /\.pdf$/i.test(decodeSafely(file.name));
    }

    // =========================================
    // ページ指定の解析 / Page range parsing
    // =========================================

    /**
     * 「1-20」「1,3,5」のような文字列をページ番号の配列に変換します。
     * @param {string} rangeText - ページ指定の文字列
     * @returns {Array<number>} ページ番号の配列
     */
    function parsePageNumbers(rangeText) {
        var pageNumbers = [];
        var tokens = String(rangeText || "").split(",");

        for (var i = 0; i < tokens.length; i++) {
            var token = tokens[i].replace(/^\s+|\s+$/g, "");

            if (token.indexOf("-") > -1) {
                var edges = token.split("-");
                var start = parseInt(edges[0], 10);
                var end = parseInt(edges[1], 10);
                if (isNaN(start) || isNaN(end)) continue;

                var firstPage = Math.min(start, end);
                var lastPage = Math.max(start, end);
                for (var j = firstPage; j <= lastPage; j++) {
                    pageNumbers.push(j);
                }
            } else {
                var singlePage = parseInt(token, 10);
                if (!isNaN(singlePage)) pageNumbers.push(singlePage);
            }
        }
        return pageNumbers;
    }

    // =========================================
    // ページ数の取得 / Page count
    // =========================================

    /**
     * 選択内容から最初の配置画像を再帰的に探します。
     * @param {Array} items - 走査するページアイテムの配列
     * @returns {PlacedItem|null} 見つかった配置画像。なければ null
     */
    function findFirstPlacedItem(items) {
        if (!items || items.length <= 0) return null;

        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) continue;

            var typeName = (item.constructor && item.constructor.name) ? item.constructor.name : "";
            if (typeName === "PlacedItem") return item;
            if (typeName === "GroupItem") {
                var found = findFirstPlacedItem(item.pageItems);
                if (found) return found;
            }
        }
        return null;
    }

    /**
     * PDF/AIファイルを走査して総ページ数を推定します。
     * @param {File} file - 読み取るPDF/AIファイル
     * @returns {number|null} 総ページ数。取得できなければ null
     */
    function readPageCountFromFile(file) {
        var countPattern = /<<\/Count\s(\d+)/;
        var linearizedPattern = /<<\/Linearized\s.+\/N\s(\d+)\/T\s.+>>/;
        var pagesTypePattern = /\/Type\/Pages/;
        var countValuePattern = /\/Count\s(\d+)/;
        var pagePatterns = [/<<\/Type\/Page\/Parent/, /\/Type\s\/Page\s/, /\/StructParents\s\d+.*\/Type\/Page>>/];

        var pageCount = 0;
        var countedPages = 0;

        try {
            if (!file.open("r")) return null;

            while (!file.eof) {
                var line = file.readln();

                /* ページ数が直接書かれていれば、その値を採用する */
                if (countPattern.test(line) || linearizedPattern.test(line)) {
                    pageCount = Number(RegExp.$1);
                    break;
                }

                /* ページを表す行を数える */
                for (var i = 0; i < pagePatterns.length; i++) {
                    if (pagePatterns[i].test(line)) {
                        countedPages++;
                        break;
                    }
                }

                /* /Type/Pages の次行に総数があれば、多いほうを採用する */
                if (pagesTypePattern.test(line)) {
                    line = file.readln();
                    if (countValuePattern.test(line)) {
                        var declaredCount = Number(RegExp.$1);
                        if (countedPages < declaredCount) countedPages = declaredCount;
                    }
                }
            }
            if (countedPages > 0) pageCount = countedPages;
        } catch (e) {
            return null;
        } finally {
            file.close();
        }

        return (pageCount > 0) ? pageCount : null;
    }

    /**
     * PDF/AIファイルの総ページ数を取得します。取得できない場合はアラートを表示します。
     * @param {File} file - 読み取るPDF/AIファイル
     * @returns {number|null} 総ページ数。取得できなければ null
     */
    function getPageCount(file) {
        var pageCount = readPageCountFromFile(file);
        if (!pageCount) {
            showAlert(LABELS.alert.pageCountFailed);
            return null;
        }
        return pageCount;
    }

    // =========================================
    // PDF読み込みの環境設定 / PDF import preferences
    // =========================================

    /**
     * PDF読み込み時のトリミング指定を環境設定に書き込みます。
     * Illustratorのバージョン差を吸収するため、複数のキーに同じ値を試します。
     * @param {number} cropValue - CROP_ART / CROP_TRIM / CROP_CROP / CROP_BLEED のいずれか
     * @returns {void}
     */
    function setPdfCropPreference(cropValue) {
        var keys = [
            "plugin/PDFImport/CropToBox",
            "plugin/PDFImport/CropTo",
            "plugin/PDFImport/CropBox",
            "plugin/PDFImport/CropToType"
        ];
        for (var i = 0; i < keys.length; i++) {
            try {
                app.preferences.setIntegerPreference(keys[i], cropValue);
            } catch (e) { }
        }
    }

    /**
     * 次に読み込むPDFのページ番号を環境設定に書き込みます。
     * @param {number} pageNumber - 読み込むページ番号（1以上）
     * @returns {void}
     */
    function setImportPageNumber(pageNumber) {
        var targetPage = parseInt(pageNumber, 10);
        if (isNaN(targetPage) || targetPage < 1) targetPage = 1;

        try {
            app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", targetPage);
        } catch (e) { }
    }

    /**
     * PDFの綴じ方向を判定します。/Direction /R2L があれば右綴じとみなします。
     * @param {File} file - 判定するPDF/AIファイル
     * @returns {boolean} 右綴じ（偶数ページが右）なら true
     */
    function isRightBoundFile(file) {
        var isRightBound = false;

        try {
            if (!file.open("r")) return false;

            var lineCount = 0;
            while (!file.eof && lineCount < BINDING_SCAN_LINE_LIMIT) {
                lineCount++;
                if (/\/Direction\s*\/R2L/.test(file.readln())) {
                    isRightBound = true;
                    break;
                }
            }
        } catch (e) {
            return false;
        } finally {
            file.close();
        }

        return isRightBound;
    }

    // =========================================
    // ページの配置 / Page placement
    // =========================================

    /**
     * 配置処理で共有する状態です。
     * @typedef {object} PlacementContext
     * @property {Document} doc - 配置先ドキュメント
     * @property {File} file - 読み込むPDF/AIファイル
     * @property {number} cropMode - トリミング指定値
     * @property {boolean} evenPageOnRight - 偶数ページを右に置くなら true
     * @property {number} activeArtboardIndex - 最初に再利用するアートボードの番号
     * @property {number} nextX - 次のアートボードの左端（pt）
     * @property {number} top - アートボードの上端（pt）
     * @property {number} artboardCount - 生成済みアートボード数
     */

    /**
     * 指定ページを一時的に配置して、実寸のページサイズを測ります。
     * @param {Document} doc - 一時配置に使うドキュメント
     * @param {File} file - 読み込むPDF/AIファイル
     * @param {number} pageNumber - 測定するページ番号
     * @param {number} cropMode - トリミング指定値
     * @returns {{width: number, height: number}} ページの幅と高さ（pt）
     */
    function measurePlacedPageSize(doc, file, pageNumber, cropMode) {
        setPdfCropPreference(cropMode);
        setImportPageNumber(pageNumber);

        var probeItem = null;
        try {
            probeItem = doc.placedItems.add();
            probeItem.file = file;
            return {
                width: probeItem.width,
                height: probeItem.height
            };
        } finally {
            /* 測定用の配置物は、元の例外を隠さないように個別に握りつぶす */
            if (probeItem) {
                try { probeItem.remove(); } catch (e) { }
            }
        }
    }

    /**
     * 配置物をアートボードの矩形でクリッピングマスクします。
     * @param {Document} doc - 対象ドキュメント
     * @param {PlacedItem} item - マスクする配置物
     * @param {Array<number>} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @returns {GroupItem} マスクを適用したグループ
     */
    function clipItemToArtboard(doc, item, artboardRect) {
        var width = Math.abs(artboardRect[2] - artboardRect[0]);
        var height = Math.abs(artboardRect[1] - artboardRect[3]);

        var clipPath = doc.activeLayer.pathItems.rectangle(artboardRect[1], artboardRect[0], width, height);
        clipPath.stroked = false;
        clipPath.filled = false;
        clipPath.clipping = true;

        var clipGroup = doc.groupItems.add();
        item.moveToEnd(clipGroup);
        clipPath.moveToBeginning(clipGroup);
        clipGroup.clipped = true;

        return clipGroup;
    }

    /**
     * 最初の1枚はアクティブなアートボードを使い、2枚目以降は新規に追加します。
     * @param {PlacementContext} context - 配置処理の状態
     * @param {Array<number>} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @returns {void}
     */
    function addOrReuseArtboard(context, artboardRect) {
        if (context.artboardCount === 0) {
            context.doc.artboards[context.activeArtboardIndex].artboardRect = artboardRect;
        } else {
            context.doc.artboards.add(artboardRect);
        }
        context.artboardCount++;
    }

    /**
     * 指定ページを配置し、アートボードの矩形でクリッピングマスクします。
     * @param {PlacementContext} context - 配置処理の状態
     * @param {number} pageNumber - 配置するページ番号
     * @param {Array<number>} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @param {Array<number>} position - 配置位置 [x, y]
     * @returns {void}
     */
    function placePageWithClip(context, pageNumber, artboardRect, position) {
        setPdfCropPreference(context.cropMode);
        setImportPageNumber(pageNumber);

        var placedItem = context.doc.placedItems.add();
        placedItem.file = context.file;
        placedItem.position = position;
        clipItemToArtboard(context.doc, placedItem, artboardRect);
    }

    /**
     * 単ページを1つのアートボードに配置します。
     * @param {PlacementContext} context - 配置処理の状態
     * @param {number} pageNumber - 配置するページ番号
     * @param {number} pageWidth - ページの幅（pt）
     * @param {number} pageHeight - ページの高さ（pt）
     * @returns {void}
     */
    function placeSinglePage(context, pageNumber, pageWidth, pageHeight) {
        var artboardRect = [context.nextX, context.top, context.nextX + pageWidth, context.top - pageHeight];

        addOrReuseArtboard(context, artboardRect);
        placePageWithClip(context, pageNumber, artboardRect, [context.nextX, context.top]);
        context.nextX += pageWidth + ARTBOARD_GAP;
    }

    /**
     * 見開きページを左右に分割し、2つのアートボードに配置します。
     * @param {PlacementContext} context - 配置処理の状態
     * @param {number} pageNumber - 配置するページ番号
     * @param {number} pageWidth - 見開き全体の幅（pt）
     * @param {number} pageHeight - ページの高さ（pt）
     * @returns {void}
     */
    function placeSpreadPage(context, pageNumber, pageWidth, pageHeight) {
        var halfWidth = pageWidth / 2;

        for (var order = 0; order < 2; order++) {
            var artboardRect = [context.nextX, context.top, context.nextX + halfWidth, context.top - pageHeight];

            /* 右綴じは右半分から、左綴じは左半分から並べる */
            var showsRightHalf = (order === 0) ? context.evenPageOnRight : !context.evenPageOnRight;
            var positionX = showsRightHalf ? (context.nextX - halfWidth) : context.nextX;

            addOrReuseArtboard(context, artboardRect);
            placePageWithClip(context, pageNumber, artboardRect, [positionX, context.top]);
            context.nextX += halfWidth + ARTBOARD_GAP;
        }
    }

    /**
     * 指定されたページを順に配置します。横長ページは見開きとして分割します。
     * @param {PlacementContext} context - 配置処理の状態
     * @param {Array<number>} pageNumbers - 配置するページ番号の配列
     * @returns {void}
     */
    function placePages(context, pageNumbers) {
        for (var i = 0; i < pageNumbers.length; i++) {
            var pageNumber = pageNumbers[i];
            if (isNaN(pageNumber) || pageNumber < 1) pageNumber = 1;

            var pageSize = measurePlacedPageSize(context.doc, context.file, pageNumber, context.cropMode);

            if (pageSize.width > pageSize.height * SPREAD_ASPECT_RATIO) {
                placeSpreadPage(context, pageNumber, pageSize.width, pageSize.height);
            } else {
                placeSinglePage(context, pageNumber, pageSize.width, pageSize.height);
            }
        }
    }

    // =========================================
    // 表示倍率の調整 / View adjustment
    // =========================================

    /**
     * 全アートボードを囲む矩形を求めます。
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array<number>|null} 矩形 [左, 上, 右, 下]。アートボードがなければ null
     */
    function getArtboardsUnionRect(doc) {
        var artboards = doc.artboards;
        if (!artboards || artboards.length === 0) return null;

        var unionRect = artboards[0].artboardRect.slice(0);
        for (var i = 1; i < artboards.length; i++) {
            var rect = artboards[i].artboardRect;
            if (rect[0] < unionRect[0]) unionRect[0] = rect[0];
            if (rect[1] > unionRect[1]) unionRect[1] = rect[1];
            if (rect[2] > unionRect[2]) unionRect[2] = rect[2];
            if (rect[3] < unionRect[3]) unionRect[3] = rect[3];
        }
        return unionRect;
    }

    /**
     * 指定した矩形が収まるように、表示位置と表示倍率を合わせます。
     * @param {View} view - 対象のビュー
     * @param {Array<number>} rect - 収める矩形 [左, 上, 右, 下]
     * @returns {void}
     */
    function fitViewToRect(view, rect) {
        var targetWidth = rect[2] - rect[0];
        var targetHeight = rect[1] - rect[3];
        if (targetWidth <= 0 || targetHeight <= 0) return;

        var bounds = view.bounds;
        var viewWidth = bounds[2] - bounds[0];
        var viewHeight = bounds[1] - bounds[3];
        var currentZoom = view.zoom;
        if (viewWidth <= 0 || viewHeight <= 0 || currentZoom <= 0) return;

        var zoomToFitWidth = currentZoom * (viewWidth / targetWidth);
        var zoomToFitHeight = currentZoom * (viewHeight / targetHeight);

        view.centerPoint = [(rect[0] + rect[2]) / 2, (rect[1] + rect[3]) / 2];
        view.zoom = Math.min(zoomToFitWidth, zoomToFitHeight) * VIEW_FIT_PADDING_RATIO;
    }

    /**
     * 全アートボードが見える表示倍率に調整します。
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function fitAllArtboardsInView(doc) {
        /* 表示の調整に失敗しても、配置結果はそのまま残す */
        try {
            app.activeDocument = doc;

            var unionRect = getArtboardsUnionRect(doc);
            var view = doc.activeView;
            if (unionRect && view) fitViewToRect(view, unionRect);
        } catch (e) { }
    }

    // =========================================
    // ダイアログの構築 / Dialog construction
    // =========================================

    /**
     * 行グループの並びと揃えを設定します。
     * @param {Group} rowGroup - 設定する行グループ
     * @returns {Group} 設定後の行グループ
     */
    function setupRow(rowGroup) {
        rowGroup.orientation = "row";
        rowGroup.alignment = ["left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        return rowGroup;
    }

    /**
     * パネルの並びと余白を設定します。
     * @param {Panel} panel - 設定するパネル
     * @returns {Panel} 設定後のパネル
     */
    function setupPanel(panel) {
        panel.orientation = "column";
        panel.alignChildren = ["left", "top"];
        panel.margins = PANEL_MARGINS;
        return panel;
    }

    /**
     * ［読み込みファイル］パネルを作成します。
     * @param {Group} parentGroup - 追加先のグループ
     * @returns {{btnSelectFile: Button, sourceNameText: StaticText}} パネル内のコントロール
     */
    function buildSourcePanel(parentGroup) {
        var sourcePanel = setupPanel(parentGroup.add("panel", undefined, getLabel(LABELS.panel.source)));

        var btnSelectFile = sourcePanel.add("button", undefined, getLabel(LABELS.button.selectFile));
        btnSelectFile.helpTip = getLabel(LABELS.tooltip.selectFile);

        var sourceNameText = sourcePanel.add("statictext", undefined, getLabel(LABELS.fieldLabel.notSelected));
        sourceNameText.characters = SOURCE_NAME_CHARS;

        return {
            btnSelectFile: btnSelectFile,
            sourceNameText: sourceNameText
        };
    }

    /**
     * ［新規ドキュメント］パネルを作成します。
     * @param {Group} parentGroup - 追加先のグループ
     * @param {Document} sourceDoc - 実行時のアクティブドキュメント
     * @returns {{rbColorCMYK: RadioButton, rbColorRGB: RadioButton}} パネル内のコントロール
     */
    function buildNewDocumentPanel(parentGroup, sourceDoc) {
        var newDocumentPanel = setupPanel(parentGroup.add("panel", undefined, getLabel(LABELS.panel.newDocument)));
        newDocumentPanel.add("statictext", undefined, getLabel(LABELS.fieldLabel.colorMode));

        var colorModeGroup = setupRow(newDocumentPanel.add("group"));
        var rbColorCMYK = colorModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.colorModeCMYK));
        var rbColorRGB = colorModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.colorModeRGB));

        /* 元ドキュメントのカラーモードを初期値にする */
        if (sourceDoc.documentColorSpace === DocumentColorSpace.RGB) rbColorRGB.value = true;
        else rbColorCMYK.value = true;

        return {
            rbColorCMYK: rbColorCMYK,
            rbColorRGB: rbColorRGB
        };
    }

    /**
     * ［アートボード］パネルを作成します。
     * @param {Group} parentGroup - 追加先のグループ
     * @returns {{etPageRange: EditText}} パネル内のコントロール
     */
    function buildArtboardPanel(parentGroup) {
        var artboardPanel = setupPanel(parentGroup.add("panel", undefined, getLabel(LABELS.panel.artboard)));

        var pageRangeGroup = setupRow(artboardPanel.add("group"));
        pageRangeGroup.add("statictext", undefined, getLabel(LABELS.fieldLabel.pageRange));

        var etPageRange = pageRangeGroup.add("edittext", undefined, "");
        etPageRange.characters = PAGE_RANGE_CHARS;
        etPageRange.helpTip = getLabel(LABELS.tooltip.pageRange);

        return {
            etPageRange: etPageRange
        };
    }

    /**
     * ［配置方法］パネルを作成します。
     * @param {Group} parentGroup - 追加先のグループ
     * @returns {{ddCropMode: DropDownList, rbEvenPageRight: RadioButton, rbEvenPageLeft: RadioButton}} パネル内のコントロール
     */
    function buildPlacementPanel(parentGroup) {
        var placementPanel = setupPanel(parentGroup.add("panel", undefined, getLabel(LABELS.panel.placement)));

        var cropNames = [];
        for (var i = 0; i < CROP_OPTIONS.length; i++) {
            cropNames.push(getLabel(CROP_OPTIONS[i].labelSet));
        }

        var ddCropMode = placementPanel.add("dropdownlist", undefined, cropNames);
        ddCropMode.minimumSize.width = CROP_LIST_MIN_WIDTH;
        ddCropMode.selection = DEFAULT_CROP_INDEX;
        ddCropMode.helpTip = getLabel(LABELS.tooltip.cropMode);
        /* トリミング指定はPDFのときだけ使うので、ファイル確定まで無効にする */
        ddCropMode.enabled = false;

        var evenPageGroup = setupRow(placementPanel.add("group"));
        evenPageGroup.helpTip = getLabel(LABELS.tooltip.evenPage);
        evenPageGroup.add("statictext", undefined, getLabel(LABELS.fieldLabel.evenPage));

        var rbEvenPageRight = evenPageGroup.add("radiobutton", undefined, getLabel(LABELS.radio.evenPageRight));
        var rbEvenPageLeft = evenPageGroup.add("radiobutton", undefined, getLabel(LABELS.radio.evenPageLeft));
        rbEvenPageRight.value = true;
        rbEvenPageRight.helpTip = getLabel(LABELS.tooltip.evenPage);
        rbEvenPageLeft.helpTip = getLabel(LABELS.tooltip.evenPage);

        return {
            ddCropMode: ddCropMode,
            rbEvenPageRight: rbEvenPageRight,
            rbEvenPageLeft: rbEvenPageLeft
        };
    }

    /**
     * ボタンエリアを作成します。OK／キャンセルはScriptUIの標準動作でダイアログを閉じます。
     * @param {Window} dialogWindow - 追加先のダイアログ
     * @returns {void}
     */
    function buildButtonRow(dialogWindow) {
        var btnRowGroup = dialogWindow.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];

        btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
    }

    /**
     * ダイアログ全体を組み立てます。
     * @param {Document} sourceDoc - 実行時のアクティブドキュメント
     * @returns {object} ダイアログと各コントロールの参照
     */
    function buildDialog(sourceDoc) {
        var dialogWindow = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        dialogWindow.alignChildren = "fill";

        var columnsGroup = dialogWindow.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];

        var leftColumnGroup = columnsGroup.add("group");
        leftColumnGroup.orientation = "column";
        leftColumnGroup.alignChildren = "fill";

        var rightColumnGroup = columnsGroup.add("group");
        rightColumnGroup.orientation = "column";
        rightColumnGroup.alignChildren = "fill";

        var sourcePanel = buildSourcePanel(leftColumnGroup);
        var newDocumentPanel = buildNewDocumentPanel(rightColumnGroup, sourceDoc);
        var artboardPanel = buildArtboardPanel(leftColumnGroup);
        var placementPanel = buildPlacementPanel(rightColumnGroup);
        buildButtonRow(dialogWindow);

        return {
            dialogWindow: dialogWindow,
            btnSelectFile: sourcePanel.btnSelectFile,
            sourceNameText: sourcePanel.sourceNameText,
            rbColorCMYK: newDocumentPanel.rbColorCMYK,
            rbColorRGB: newDocumentPanel.rbColorRGB,
            etPageRange: artboardPanel.etPageRange,
            ddCropMode: placementPanel.ddCropMode,
            rbEvenPageRight: placementPanel.rbEvenPageRight,
            rbEvenPageLeft: placementPanel.rbEvenPageLeft
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    function main() {

        if (app.documents.length === 0) {
            showAlert(LABELS.alert.needDocument);
            return;
        }

        var sourceDoc = app.activeDocument;

        /* 現在の読み込み対象ファイル。未選択時は null。 */
        var sourceFile = null;

        var dialogUI = buildDialog(sourceDoc);
        var dialogWindow = dialogUI.dialogWindow;
        var btnSelectFile = dialogUI.btnSelectFile;
        var sourceNameText = dialogUI.sourceNameText;
        var rbColorRGB = dialogUI.rbColorRGB;
        var etPageRange = dialogUI.etPageRange;
        var ddCropMode = dialogUI.ddCropMode;
        var rbEvenPageRight = dialogUI.rbEvenPageRight;
        var rbEvenPageLeft = dialogUI.rbEvenPageLeft;

        // ------------------------
        // UIの更新
        // ------------------------

        /**
         * 読み込み対象ファイルを差し替え、ファイル名・トリミング指定・綴じ方向を更新します。
         * @param {File} file - 読み込み対象のPDF/AIファイル
         * @returns {void}
         */
        function updateSourceFile(file) {
            sourceFile = file || null;

            /* トリミング指定はPDFのときのみ有効 */
            ddCropMode.enabled = isPdfFile(sourceFile);

            /* PDFの綴じ方向から偶数ページの位置を自動設定 */
            if (sourceFile) {
                if (isRightBoundFile(sourceFile)) rbEvenPageRight.value = true;
                else rbEvenPageLeft.value = true;
            }

            sourceNameText.text = sourceFile ? decodeSafely(sourceFile.name) : getLabel(LABELS.fieldLabel.notSelected);
            sourceNameText.helpTip = sourceFile ? decodeSafely(sourceFile.fsName) : "";
        }

        /**
         * ページ指定欄に、取得した総ページ数から初期値を入れます。
         * @param {number|null} pageCount - 総ページ数
         * @returns {void}
         */
        function updatePageRangeField(pageCount) {
            etPageRange.text = pageCount ? ("1-" + pageCount) : "";
        }

        /**
         * ［ファイル指定］で選んだファイルをダイアログに反映します。
         * @param {File} file - 選択したファイル
         * @returns {void}
         */
        function loadSourceFile(file) {
            if (!isPdfOrAiFile(file)) {
                showAlert(LABELS.alert.pickPdfAi);
                updatePageRangeField(null);
                return;
            }

            var pageCount = getPageCount(file);
            updateSourceFile(file);
            updatePageRangeField(pageCount);
        }

        /**
         * 選択中の配置画像がPDF/AIなら、そのリンク先をダイアログに反映します。
         * @returns {void}
         */
        function loadSourceFromSelection() {
            var placedItem = findFirstPlacedItem(sourceDoc.selection);
            if (!placedItem) return;

            if (!placedItem.file) {
                showAlert(LABELS.alert.linkUnknown);
                return;
            }
            if (!isPdfOrAiFile(placedItem.file)) return;

            var pageCount = getPageCount(placedItem.file);
            updateSourceFile(placedItem.file);
            updatePageRangeField(pageCount);
        }

        // ------------------------
        // ダイアログの設定値
        // ------------------------

        /**
         * 新規ドキュメントのカラーモードを返します。
         * @returns {DocumentColorSpace} 選択されたカラーモード
         */
        function getSelectedColorSpace() {
            return rbColorRGB.value ? DocumentColorSpace.RGB : DocumentColorSpace.CMYK;
        }

        /**
         * 選択されたトリミング指定値を返します。
         * @returns {number} CROP_ART / CROP_TRIM / CROP_CROP / CROP_BLEED のいずれか
         */
        function getSelectedCropMode() {
            var index = ddCropMode.selection ? ddCropMode.selection.index : DEFAULT_CROP_INDEX;
            var option = CROP_OPTIONS[index];
            return option ? option.value : CROP_CROP;
        }

        // ------------------------
        // 実行処理
        // ------------------------

        /**
         * 新規ドキュメントの初期サイズを、最初に配置するページから求めます。
         * @param {number} cropMode - トリミング指定値
         * @returns {{width: number, height: number}} ドキュメントの幅と高さ（pt）
         */
        function getOutputDocSize(cropMode) {
            var pageNumbers = parsePageNumbers(etPageRange.text);
            var firstPage = (pageNumbers.length > 0) ? pageNumbers[0] : 1;
            if (firstPage < 1) firstPage = 1;

            /* 測定できないときは元ドキュメントのサイズで作成する */
            try {
                return measurePlacedPageSize(sourceDoc, sourceFile, firstPage, cropMode);
            } catch (e) {
                return {
                    width: sourceDoc.width,
                    height: sourceDoc.height
                };
            }
        }

        /**
         * 配置先の新規ドキュメントを作成します。
         * @param {number} cropMode - トリミング指定値
         * @returns {Document} 作成したドキュメント
         */
        function createOutputDoc(cropMode) {
            var outputSize = getOutputDocSize(cropMode);
            var outputDoc = app.documents.add(getSelectedColorSpace(), outputSize.width, outputSize.height);

            /* ラスタライズ効果解像度の設定に失敗しても配置は続行する */
            try {
                var rasterSettings = outputDoc.rasterEffectSettings;
                rasterSettings.resolution = RASTER_EFFECTS_RESOLUTION;
                outputDoc.rasterEffectSettings = rasterSettings;
            } catch (e) { }

            app.activeDocument = outputDoc;
            return outputDoc;
        }

        /**
         * 新規ドキュメントを作成し、指定ページを個別のアートボードに配置します。
         * @param {Array<number>} pageNumbers - 配置するページ番号の配列
         * @param {number} cropMode - トリミング指定値
         * @param {boolean} evenPageOnRight - 偶数ページを右に置くなら true
         * @returns {void}
         */
        function importPages(pageNumbers, cropMode, evenPageOnRight) {
            var outputDoc = createOutputDoc(cropMode);
            var activeArtboardIndex = outputDoc.artboards.getActiveArtboardIndex();
            var baseRect = outputDoc.artboards[activeArtboardIndex].artboardRect;

            var context = {
                doc: outputDoc,
                file: sourceFile,
                cropMode: cropMode,
                evenPageOnRight: evenPageOnRight,
                activeArtboardIndex: activeArtboardIndex,
                nextX: baseRect[0],
                top: baseRect[1],
                artboardCount: 0
            };

            var completed = false;
            try {
                placePages(context, pageNumbers);
                completed = true;
            } catch (e) {
                showAlert(LABELS.alert.placeFailed, e);
            } finally {
                /* 読み込みページ番号の環境設定を既定に戻す */
                setImportPageNumber(1);
            }

            if (completed) fitAllArtboardsInView(outputDoc);
        }

        // ------------------------
        // イベントと初期表示
        // ------------------------

        btnSelectFile.onClick = function () {
            var file = File.openDialog(getLabel(LABELS.dialog.pickFile), getLabel(LABELS.dialog.pickFilter));
            if (file) loadSourceFile(file);
        };

        loadSourceFromSelection();

        if (dialogWindow.show() !== 1) return;

        if (!sourceFile) {
            showAlert(LABELS.alert.needFile);
            return;
        }

        var pageNumbers = parsePageNumbers(etPageRange.text);
        if (pageNumbers.length === 0) pageNumbers = [1];

        importPages(pageNumbers, getSelectedCropMode(), rbEvenPageRight.value);
    }

    main();

})();
