#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

PDF/AI ファイルを指定したページ範囲で読み込み、新規ドキュメント上に各ページを個別のアートボードとして配置します。
横長ページは見開きとして自動判定し、左右2つのアートボードに分割します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PDFAISpreadImporter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n5514d9f2c5f8

### Overview

Imports a PDF/AI file over a given page range and places each page on its own artboard in a new document.
Landscape pages are detected as spreads and split into two artboards, left and right.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PDFAISpreadImporter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PDFAISpreadImporter";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-25";                   /* 更新日 / last updated */

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

    /* トリミング指定の初期選択（CROP_OPTIONS の番号。2 = 仕上がり） / Default crop option (index into CROP_OPTIONS; 2 = Trim) */
    var DEFAULT_CROP_INDEX = 2;

    // =========================================
    // PDF読み込みの環境設定 / PDF import preferences
    // =========================================

    /* PDF読み込み時のトリミング指定値（CropTo の値。4 = メディア、5 = バウンディングボックス）
       Crop-to values written to CropTo (4 = Media, 5 = Bounding box) */
    var CROP_ART = 0;
    var CROP_CROP = 1;
    var CROP_TRIM = 2;
    var CROP_BLEED = 3;

    /* トリミング指定の環境設定キー / Preference key for the crop-to box */
    var PDF_CROP_PREF_KEY = "plugin/PDFImport/CropTo";

    /* 次に読み込むページ番号の環境設定キー / Preference key for the page number to import next */
    var PDF_PAGE_NUMBER_PREF_KEY = "plugin/PDFImport/PageNumber";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10];
    var BUTTON_ROW_TOP_MARGIN = 10;
    var SOURCE_NAME_CHARS = 16;
    var PAGE_RANGE_CHARS = 10;
    var CROP_LIST_MIN_WIDTH = 160;

    /**
     * 行グループの並びと揃えを設定します。
     * @param {Group} rowGroup - 設定する行グループ
     * @returns {Group} 設定後の行グループ
     */
    function setupRowGroup(rowGroup) {
        rowGroup.orientation = "row";
        rowGroup.alignment = ["left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        return rowGroup;
    }

    /**
     * パネルの並びと余白を設定します。
     * @param {Panel} targetPanel - 設定するパネル
     * @returns {Panel} 設定後のパネル
     */
    function setupPanel(targetPanel) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["left", "top"];
        targetPanel.margins = PANEL_MARGINS;
        return targetPanel;
    }

    // =========================================
    // ローカライズ / Localization
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
            pickFile: { ja: "PDF/AIファイルを選択してください", en: "Select a PDF/AI file" },
            pickFilter: { ja: "PDF/AI:*.pdf;*.ai", en: "PDF/AI:*.pdf;*.ai" }
        },

        panel: {
            source: { ja: "読み込みファイル", en: "Source File" },
            pages: { ja: "ページ", en: "Pages" },
            placement: { ja: "配置方法", en: "Placement" },
            newDocument: { ja: "新規ドキュメント", en: "New Document" }
        },

        fieldLabel: {
            pageRange: { ja: "範囲", en: "Range" },
            cropMode: { ja: "トリミング", en: "Crop to" },
            evenPage: { ja: "偶数ページ", en: "Even pages" },
            colorMode: { ja: "カラーモード", en: "Color mode" }
        },

        radio: {
            evenPageRight: { ja: "右", en: "Right" },
            evenPageLeft: { ja: "左", en: "Left" },
            colorModeCMYK: { ja: "CMYK", en: "CMYK" },
            colorModeRGB: { ja: "RGB", en: "RGB" }
        },

        dropdown: {
            cropArt: { ja: "アート", en: "Art" },
            cropCrop: { ja: "トリミング", en: "Crop" },
            cropTrim: { ja: "仕上がり", en: "Trim" },
            cropBleed: { ja: "裁ち落とし", en: "Bleed" }
        },

        button: {
            selectFile: { ja: "ファイルを選択...", en: "Choose File..." },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },

        tooltip: {
            selectFile: {
                ja: "読み込むPDF/AIファイルを選びます。選択中の配置画像があれば、起動時にそのリンク先を読み込みます。",
                en: "Choose the PDF/AI file to import. A selected placed image is picked up on startup."
            },
            pageRange: {
                ja: "読み込むページを「1-10」「1,3,5」のように指定します。ファイルを選ぶと全ページが入ります。空欄のときは1ページ目だけを配置します。",
                en: "Pages to import, written as 1-10 or 1,3,5. Filled with all pages when a file is chosen. Leave it empty to place only the first page."
            },
            cropMode: {
                ja: "PDFのどの領域を基準に配置するかを選びます。AIファイルでは使用しません。",
                en: "Which PDF box the pages are placed from. Not used for AI files."
            },
            evenPage: {
                ja: "見開きを左右に分割したとき、偶数ページを置く側です。PDFの綴じ方向から自動で設定します。",
                en: "Which side the even pages go to when a spread is split. Detected from the PDF binding direction."
            },
            colorMode: {
                ja: "作成する新規ドキュメントのカラーモードです。初期値は現在のドキュメントに合わせます。",
                en: "Color mode of the new document. Defaults to that of the current document."
            }
        },

        alert: {
            needDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document first." },
            needFile: { ja: "先に［ファイルを選択...］で読み込むファイルを選んでください。", en: "Please choose a source file first." },
            pickPdfAi: { ja: "PDFまたはAIファイルを選択してください。", en: "Please select a PDF or AI file." },
            linkUnknown: { ja: "選択中の配置画像のリンク先が見つかりませんでした。", en: "Could not find the linked file of the selected placed image." },
            pageCountFailed: { ja: "PDF/AIファイルのページ数を取得できませんでした。", en: "Could not determine the page count of the PDF/AI file." },
            placeFailed: { ja: "配置中にエラーが発生しました。", en: "An error occurred while placing the pages." },
            canvasFull: {
                ja: "アートボードがキャンバスに収まりません。ページ範囲を分けて実行してください。",
                en: "The artboards do not fit on the canvas. Split the page range and run again."
            },
            errorDetails: { ja: "詳細：", en: "Details:" }
        },

        fallbackName: {
            noSourceFile: { ja: "未選択", en: "No file selected" }
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

    /**
     * 項目名にコロンを付けて返します（日本語は全角、英語は半角）。
     * @param {object} labelSet - { ja: string, en: string } 形式のラベル定義
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    /* トリミング指定の選択肢。ドロップダウンの並び順と一致させます（Illustratorの配置オプションと同じ順）。 */
    var CROP_OPTIONS = [
        { value: CROP_ART, labelSet: LABELS.dropdown.cropArt },
        { value: CROP_CROP, labelSet: LABELS.dropdown.cropCrop },
        { value: CROP_TRIM, labelSet: LABELS.dropdown.cropTrim },
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
        return String(e.message ? e.message : e);
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
    // ファイルの判定と走査 / File checks and scanning
    // =========================================

    /**
     * URLエンコードされた文字列を復号します。復号できない場合は元の文字列を返します。
     * @param {string} encodedText - 復号する文字列
     * @returns {string} 復号後の文字列
     */
    function decodeSafely(encodedText) {
        var rawText = String(encodedText || "");
        /* 不正な % 並びは URIError になる / Malformed % sequences throw URIError */
        try {
            return decodeURIComponent(rawText);
        } catch (e) {
            return rawText;
        }
    }

    /**
     * ファイル名の拡張子がパターンに合うかを判定します。
     * @param {File} targetFile - 判定するファイル
     * @param {RegExp} extensionPattern - 拡張子のパターン
     * @returns {boolean} 合えば true
     */
    function hasFileExtension(targetFile, extensionPattern) {
        return !!targetFile && extensionPattern.test(decodeSafely(targetFile.name));
    }

    /**
     * PDFまたはAIファイルかどうかを拡張子で判定します。
     * @param {File} targetFile - 判定するファイル
     * @returns {boolean} PDFまたはAIなら true
     */
    function isPdfOrAiFile(targetFile) {
        return hasFileExtension(targetFile, /\.(?:pdf|ai)$/i);
    }

    /**
     * PDFファイルかどうかを拡張子で判定します。
     * @param {File} targetFile - 判定するファイル
     * @returns {boolean} PDFなら true
     */
    function isPdfFile(targetFile) {
        return hasFileExtension(targetFile, /\.pdf$/i);
    }

    /**
     * ファイルを1行ずつ読み、コールバックが true を返すか上限に達したら止めます。
     * @param {File} targetFile - 読み取るファイル
     * @param {number} lineLimit - 読む行数の上限。0 なら末尾まで
     * @param {function(string): boolean} onLine - 1行ごとに呼ぶ関数。true で打ち切り
     * @returns {boolean} 最後まで読み取れたら true、開けないか読み取りに失敗したら false
     */
    function scanFileLines(targetFile, lineLimit, onLine) {
        if (!targetFile.open("r")) return false;

        /* バイナリを含むPDFは読み取り中に例外になることがある / Binary PDF content can throw while reading */
        try {
            var lineCount = 0;
            while (!targetFile.eof && (lineLimit === 0 || lineCount < lineLimit)) {
                lineCount++;
                if (onLine(targetFile.readln())) break;
            }
            return true;
        } catch (e) {
            return false;
        } finally {
            targetFile.close();
        }
    }

    /**
     * PDF/AIファイルを走査して総ページ数を推定します。
     * @param {File} sourceFile - 読み取るPDF/AIファイル
     * @returns {number|null} 総ページ数。取得できなければ null
     */
    function readPageCountFromFile(sourceFile) {
        var countPattern = /<<\/Count\s(\d+)/;
        var linearizedPattern = /<<\/Linearized\s.+\/N\s(\d+)\/T\s.+>>/;
        var pagesTypePattern = /\/Type\/Pages/;
        var countValuePattern = /\/Count\s(\d+)/;
        var pagePatterns = [/<<\/Type\/Page\/Parent/, /\/Type\s\/Page\s/, /\/StructParents\s\d+.*\/Type\/Page>>/];

        var pageCount = 0;
        var countedPages = 0;
        var expectsPagesCount = false;

        var readOk = scanFileLines(sourceFile, 0, function (lineText) {
            /* /Type/Pages の次行に総数があれば、多いほうを採用する */
            if (expectsPagesCount) {
                expectsPagesCount = false;
                if (countValuePattern.test(lineText)) {
                    var declaredCount = Number(RegExp.$1);
                    if (countedPages < declaredCount) countedPages = declaredCount;
                }
                return false;
            }

            /* ページ数が直接書かれていれば、その値を採用する */
            if (countPattern.test(lineText) || linearizedPattern.test(lineText)) {
                pageCount = Number(RegExp.$1);
                return true;
            }

            /* ページを表す行を数える */
            for (var i = 0; i < pagePatterns.length; i++) {
                if (pagePatterns[i].test(lineText)) {
                    countedPages++;
                    break;
                }
            }

            if (pagesTypePattern.test(lineText)) expectsPagesCount = true;
            return false;
        });
        if (!readOk) return null;

        if (countedPages > 0) pageCount = countedPages;
        return (pageCount > 0) ? pageCount : null;
    }

    /**
     * PDFの綴じ方向を判定します。/Direction /R2L があれば右綴じとみなします。
     * @param {File} sourceFile - 判定するPDF/AIファイル
     * @returns {boolean} 右綴じ（偶数ページが右）なら true
     */
    function isRightBoundFile(sourceFile) {
        var isRightBound = false;
        var readOk = scanFileLines(sourceFile, BINDING_SCAN_LINE_LIMIT, function (lineText) {
            isRightBound = /\/Direction\s*\/R2L/.test(lineText);
            return isRightBound;
        });
        return readOk && isRightBound;
    }

    /**
     * 選択内容から最初の配置画像を再帰的に探します。
     * @param {Array} pageItems - 走査するページアイテムの配列
     * @returns {PlacedItem|null} 見つかった配置画像。なければ null
     */
    function findFirstPlacedItem(pageItems) {
        if (!pageItems) return null;

        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (!pageItem) continue;

            if (pageItem.typename === "PlacedItem") return pageItem;
            if (pageItem.typename === "GroupItem") {
                var nestedPlacedItem = findFirstPlacedItem(pageItem.pageItems);
                if (nestedPlacedItem) return nestedPlacedItem;
            }
        }
        return null;
    }

    // =========================================
    // ページ指定の解析 / Page range parsing
    // =========================================

    /**
     * 「1-20」「1,3,5」のような文字列をページ番号の配列に変換します。1未満は1に寄せます。
     * @param {string} rangeText - ページ指定の文字列
     * @returns {number[]} ページ番号の配列
     */
    function parsePageNumbers(rangeText) {
        var pageNumbers = [];
        var rangeTokens = String(rangeText || "").split(",");

        for (var i = 0; i < rangeTokens.length; i++) {
            var rangeToken = rangeTokens[i].replace(/^\s+|\s+$/g, "");

            if (rangeToken.indexOf("-") > -1) {
                var rangeEdges = rangeToken.split("-");
                var rangeStart = parseInt(rangeEdges[0], 10);
                var rangeEnd = parseInt(rangeEdges[1], 10);
                if (isNaN(rangeStart) || isNaN(rangeEnd)) continue;

                var lastPage = Math.max(rangeStart, rangeEnd);
                for (var j = Math.min(rangeStart, rangeEnd); j <= lastPage; j++) {
                    pageNumbers.push(Math.max(1, j));
                }
            } else {
                var singlePage = parseInt(rangeToken, 10);
                if (!isNaN(singlePage)) pageNumbers.push(Math.max(1, singlePage));
            }
        }
        return pageNumbers;
    }

    // =========================================
    // ページの配置 / Page placement
    // =========================================

    /**
     * 次に読み込むPDFのトリミング指定とページ番号を環境設定に書き込みます。
     * @param {number} cropMode - CROP_ART / CROP_TRIM / CROP_CROP / CROP_BLEED のいずれか
     * @param {number} pageNumber - 読み込むページ番号（1以上）
     * @returns {void}
     */
    function setPdfImportPreferences(cropMode, pageNumber) {
        app.preferences.setIntegerPreference(PDF_CROP_PREF_KEY, cropMode);
        app.preferences.setIntegerPreference(PDF_PAGE_NUMBER_PREF_KEY, pageNumber);
    }

    /**
     * 配置処理で共有する状態です。
     * @typedef {object} PlacementState
     * @property {Document} doc - 配置先ドキュメント
     * @property {File} sourceFile - 読み込むPDF/AIファイル
     * @property {number} cropMode - トリミング指定値
     * @property {boolean} evenPageOnRight - 偶数ページを右に置くなら true
     * @property {number} activeArtboardIndex - 最初に再利用するアートボードの番号
     * @property {number} nextX - 次のアートボードの左端（pt）
     * @property {number} top - 今の行の上端（pt）
     * @property {number} rowLeft - 行の左端（pt）。折り返すとここに戻る
     * @property {number} rowHeight - 今の行でいちばん高いページの高さ（pt）
     * @property {number} canvasRight - キャンバスの右端（pt）
     * @property {number} canvasBottom - キャンバスの下端（pt）
     * @property {number} artboardCount - 生成済みアートボード数
     */

    /* キャンバスの一辺の半分（pt）。アートボードはこの外に置けない（Error 'AOoC'）
       Half the canvas side (pt); artboards outside it fail with 'AOoC' */
    var CANVAS_HALF_SIZE = 16383 / 2;

    /**
     * 次に置く幅がキャンバスの右端を越えるなら、次の行へ折り返します。
     * 下端も越えるときは、配置を続けられないので例外にします。
     * @param {PlacementState} placementState - 配置処理の状態
     * @param {number} neededWidth - これから並べる幅（pt）。見開きは左右2枚分
     * @param {number} pageHeight - ページの高さ（pt）
     * @returns {void}
     */
    function reserveRowSpace(placementState, neededWidth, pageHeight) {
        var rowHasItems = placementState.nextX > placementState.rowLeft;
        if (rowHasItems && placementState.nextX + neededWidth > placementState.canvasRight) {
            placementState.top -= placementState.rowHeight + ARTBOARD_GAP;
            placementState.nextX = placementState.rowLeft;
            placementState.rowHeight = 0;
        }
        if (placementState.top - pageHeight < placementState.canvasBottom) {
            throw new Error(getLabel(LABELS.alert.canvasFull));
        }
        placementState.rowHeight = Math.max(placementState.rowHeight, pageHeight);
    }

    /**
     * 指定ページを一時的に配置して、実寸のページサイズを測ります。
     * @param {Document} doc - 一時配置に使うドキュメント
     * @param {File} sourceFile - 読み込むPDF/AIファイル
     * @param {number} pageNumber - 測定するページ番号
     * @param {number} cropMode - トリミング指定値
     * @returns {{width: number, height: number}} ページの幅と高さ（pt）
     */
    function measurePlacedPageSize(doc, sourceFile, pageNumber, cropMode) {
        setPdfImportPreferences(cropMode, pageNumber);

        var probeItem = null;
        try {
            probeItem = doc.placedItems.add();
            probeItem.file = sourceFile;
            return { width: probeItem.width, height: probeItem.height };
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
     * @param {PlacedItem} placedItem - マスクする配置物
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @returns {GroupItem} マスクを適用したグループ
     */
    function clipItemToArtboard(doc, placedItem, artboardRect) {
        var clipWidth = Math.abs(artboardRect[2] - artboardRect[0]);
        var clipHeight = Math.abs(artboardRect[1] - artboardRect[3]);

        var clipPath = doc.activeLayer.pathItems.rectangle(artboardRect[1], artboardRect[0], clipWidth, clipHeight);
        clipPath.stroked = false;
        clipPath.filled = false;
        clipPath.clipping = true;

        var clipGroup = doc.groupItems.add();
        placedItem.moveToEnd(clipGroup);
        clipPath.moveToBeginning(clipGroup);
        clipGroup.clipped = true;

        return clipGroup;
    }

    /**
     * 1枚分のアートボードを用意し、ページを配置してクリッピングマスクします。
     * 最初の1枚はアクティブなアートボードを使い、2枚目以降は新規に追加します。
     * @param {PlacementState} placementState - 配置処理の状態
     * @param {number} pageNumber - 配置するページ番号
     * @param {number} artboardWidth - アートボードの幅（pt）
     * @param {number} pageHeight - ページの高さ（pt）
     * @param {number} placeX - 配置物の左端（pt）
     * @returns {void}
     */
    function placeOnNextArtboard(placementState, pageNumber, artboardWidth, pageHeight, placeX) {
        var doc = placementState.doc;
        var left = placementState.nextX;
        var top = placementState.top;
        var artboardRect = [left, top, left + artboardWidth, top - pageHeight];

        if (placementState.artboardCount === 0) {
            doc.artboards[placementState.activeArtboardIndex].artboardRect = artboardRect;
        } else {
            doc.artboards.add(artboardRect);
        }
        placementState.artboardCount++;

        setPdfImportPreferences(placementState.cropMode, pageNumber);
        var placedItem = doc.placedItems.add();
        placedItem.file = placementState.sourceFile;
        placedItem.position = [placeX, top];
        clipItemToArtboard(doc, placedItem, artboardRect);

        placementState.nextX += artboardWidth + ARTBOARD_GAP;
    }

    /**
     * 見開きページを左右に分割し、2つのアートボードに配置します。
     * @param {PlacementState} placementState - 配置処理の状態
     * @param {number} pageNumber - 配置するページ番号
     * @param {number} pageWidth - 見開き全体の幅（pt）
     * @param {number} pageHeight - ページの高さ（pt）
     * @returns {void}
     */
    function placeSpreadPage(placementState, pageNumber, pageWidth, pageHeight) {
        var halfWidth = pageWidth / 2;

        for (var halfIndex = 0; halfIndex < 2; halfIndex++) {
            /* 右綴じは右半分から、左綴じは左半分から並べる */
            var showsRightHalf = (halfIndex === 0) ? placementState.evenPageOnRight : !placementState.evenPageOnRight;
            var placeX = showsRightHalf ? (placementState.nextX - halfWidth) : placementState.nextX;

            placeOnNextArtboard(placementState, pageNumber, halfWidth, pageHeight, placeX);
        }
    }

    /**
     * 指定されたページを順に配置します。横長ページは見開きとして分割します。
     * @param {PlacementState} placementState - 配置処理の状態
     * @param {number[]} pageNumbers - 配置するページ番号の配列
     * @returns {void}
     */
    function placePages(placementState, pageNumbers) {
        for (var i = 0; i < pageNumbers.length; i++) {
            var pageNumber = pageNumbers[i];
            var pageSize = measurePlacedPageSize(placementState.doc, placementState.sourceFile, pageNumber, placementState.cropMode);

            var isSpread = pageSize.width > pageSize.height * SPREAD_ASPECT_RATIO;
            /* 見開きは左右2枚を同じ行に置く / Keep both halves of a spread on one row */
            reserveRowSpace(placementState, isSpread ? pageSize.width + ARTBOARD_GAP : pageSize.width, pageSize.height);

            if (isSpread) {
                placeSpreadPage(placementState, pageNumber, pageSize.width, pageSize.height);
            } else {
                placeOnNextArtboard(placementState, pageNumber, pageSize.width, pageSize.height, placementState.nextX);
            }
        }
    }

    // =========================================
    // 表示倍率の調整 / View adjustment
    // =========================================

    /**
     * 全アートボードを囲む矩形を求めます。
     * @param {Document} doc - 対象ドキュメント
     * @returns {number[]|null} 矩形 [左, 上, 右, 下]。アートボードがなければ null
     */
    function getArtboardsUnionRect(doc) {
        var artboards = doc.artboards;
        if (artboards.length === 0) return null;

        var unionRect = artboards[0].artboardRect.slice(0);
        for (var i = 1; i < artboards.length; i++) {
            var artboardRect = artboards[i].artboardRect;
            if (artboardRect[0] < unionRect[0]) unionRect[0] = artboardRect[0];
            if (artboardRect[1] > unionRect[1]) unionRect[1] = artboardRect[1];
            if (artboardRect[2] > unionRect[2]) unionRect[2] = artboardRect[2];
            if (artboardRect[3] < unionRect[3]) unionRect[3] = artboardRect[3];
        }
        return unionRect;
    }

    /**
     * 指定した矩形が収まるように、表示位置と表示倍率を合わせます。
     * @param {View} targetView - 対象のビュー
     * @param {number[]} targetRect - 収める矩形 [左, 上, 右, 下]
     * @returns {void}
     */
    function fitViewToRect(targetView, targetRect) {
        var targetWidth = targetRect[2] - targetRect[0];
        var targetHeight = targetRect[1] - targetRect[3];
        if (targetWidth <= 0 || targetHeight <= 0) return;

        var viewBounds = targetView.bounds;
        var viewWidth = viewBounds[2] - viewBounds[0];
        var viewHeight = viewBounds[1] - viewBounds[3];
        var currentZoom = targetView.zoom;
        if (viewWidth <= 0 || viewHeight <= 0 || currentZoom <= 0) return;

        var zoomToFitWidth = currentZoom * (viewWidth / targetWidth);
        var zoomToFitHeight = currentZoom * (viewHeight / targetHeight);

        targetView.centerPoint = [(targetRect[0] + targetRect[2]) / 2, (targetRect[1] + targetRect[3]) / 2];
        targetView.zoom = Math.min(zoomToFitWidth, zoomToFitHeight) * VIEW_FIT_PADDING_RATIO;
    }

    /**
     * 全アートボードが見える表示倍率に調整します。
     * @param {Document} doc - 対象ドキュメント（アクティブであること）
     * @returns {void}
     */
    function fitAllArtboardsInView(doc) {
        /* 表示の調整に失敗しても、配置結果はそのまま残す */
        try {
            var unionRect = getArtboardsUnionRect(doc);
            var activeView = doc.activeView;
            if (unionRect && activeView) fitViewToRect(activeView, unionRect);
        } catch (e) { }
    }

    // =========================================
    // 新規ドキュメントへの読み込み / Import into a new document
    // =========================================

    /**
     * ダイアログで決めた読み込み条件です。
     * @typedef {object} ImportSettings
     * @property {File} sourceFile - 読み込むPDF/AIファイル
     * @property {number[]} pageNumbers - 配置するページ番号の配列（1件以上）
     * @property {number} cropMode - トリミング指定値
     * @property {DocumentColorSpace} colorSpace - 新規ドキュメントのカラーモード
     * @property {boolean} evenPageOnRight - 偶数ページを右に置くなら true
     */

    /**
     * 配置先の新規ドキュメントを作成してアクティブにします。
     * 初期サイズは最初に配置するページに合わせ、測れなければ元ドキュメントのサイズにします。
     * @param {Document} sourceDoc - 実行時のアクティブドキュメント
     * @param {ImportSettings} importSettings - 読み込み条件
     * @returns {Document} 作成したドキュメント
     */
    function createOutputDoc(sourceDoc, importSettings) {
        var outputSize;
        /* 読めないPDFは placedItems.add() の file 代入で例外になる / An unreadable PDF throws on file assignment */
        try {
            outputSize = measurePlacedPageSize(sourceDoc, importSettings.sourceFile, importSettings.pageNumbers[0], importSettings.cropMode);
        } catch (e) {
            outputSize = { width: sourceDoc.width, height: sourceDoc.height };
        }

        var outputDoc = app.documents.add(importSettings.colorSpace, outputSize.width, outputSize.height);

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
     * @param {Document} sourceDoc - 実行時のアクティブドキュメント
     * @param {ImportSettings} importSettings - 読み込み条件
     * @returns {void}
     */
    function importPages(sourceDoc, importSettings) {
        var outputDoc = createOutputDoc(sourceDoc, importSettings);
        var activeArtboardIndex = outputDoc.artboards.getActiveArtboardIndex();
        var baseRect = outputDoc.artboards[activeArtboardIndex].artboardRect;

        var placementState = {
            doc: outputDoc,
            sourceFile: importSettings.sourceFile,
            cropMode: importSettings.cropMode,
            evenPageOnRight: importSettings.evenPageOnRight,
            activeArtboardIndex: activeArtboardIndex,
            nextX: baseRect[0],
            top: baseRect[1],
            rowLeft: baseRect[0],
            rowHeight: 0,
            /* 新規ドキュメントの最初のアートボードはキャンバスの中央にある / The first artboard sits at the canvas center */
            canvasRight: (baseRect[0] + baseRect[2]) / 2 + CANVAS_HALF_SIZE,
            canvasBottom: (baseRect[1] + baseRect[3]) / 2 - CANVAS_HALF_SIZE,
            artboardCount: 0
        };

        var placedAll = false;
        try {
            placePages(placementState, importSettings.pageNumbers);
            placedAll = true;
        } catch (e) {
            showAlert(LABELS.alert.placeFailed, e);
        } finally {
            /* 読み込みページ番号の環境設定を既定に戻す */
            app.preferences.setIntegerPreference(PDF_PAGE_NUMBER_PREF_KEY, 1);
        }

        if (placedAll) fitAllArtboardsInView(outputDoc);
    }

    // =========================================
    // ダイアログの構築 / Dialog construction
    // =========================================

    /**
     * ［読み込みファイル］パネルを作成します。
     * @param {Group} parentGroup - 追加先のグループ
     * @returns {{btnSelectFile: Button, sourceNameText: StaticText}} パネル内のコントロール
     */
    function buildSourcePanel(parentGroup) {
        var sourcePanel = setupPanel(parentGroup.add("panel", undefined, getLabel(LABELS.panel.source)));

        var btnSelectFile = sourcePanel.add("button", undefined, getLabel(LABELS.button.selectFile));
        btnSelectFile.helpTip = getLabel(LABELS.tooltip.selectFile);

        var sourceNameText = sourcePanel.add("statictext", undefined, getLabel(LABELS.fallbackName.noSourceFile));
        sourceNameText.characters = SOURCE_NAME_CHARS;

        return { btnSelectFile: btnSelectFile, sourceNameText: sourceNameText };
    }

    /**
     * ［新規ドキュメント］パネルを作成します。
     * @param {Group} parentGroup - 追加先のグループ
     * @param {Document} sourceDoc - 実行時のアクティブドキュメント
     * @returns {{rbColorRGB: RadioButton}} パネル内のコントロール
     */
    function buildNewDocumentPanel(parentGroup, sourceDoc) {
        var newDocumentPanel = setupPanel(parentGroup.add("panel", undefined, getLabel(LABELS.panel.newDocument)));
        var colorModeLabel = newDocumentPanel.add("statictext", undefined, labelText(LABELS.fieldLabel.colorMode));

        var colorModeGroup = setupRowGroup(newDocumentPanel.add("group"));
        var rbColorCMYK = colorModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.colorModeCMYK));
        var rbColorRGB = colorModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.colorModeRGB));
        colorModeLabel.helpTip = rbColorCMYK.helpTip = rbColorRGB.helpTip = getLabel(LABELS.tooltip.colorMode);

        /* 元ドキュメントのカラーモードを初期値にする */
        if (sourceDoc.documentColorSpace === DocumentColorSpace.RGB) rbColorRGB.value = true;
        else rbColorCMYK.value = true;

        return { rbColorRGB: rbColorRGB };
    }

    /**
     * ［ページ］パネルを作成します。
     * @param {Group} parentGroup - 追加先のグループ
     * @returns {{etPageRange: EditText}} パネル内のコントロール
     */
    function buildPagesPanel(parentGroup) {
        var pagesPanel = setupPanel(parentGroup.add("panel", undefined, getLabel(LABELS.panel.pages)));

        var pageRangeGroup = setupRowGroup(pagesPanel.add("group"));
        pageRangeGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.pageRange));

        var etPageRange = pageRangeGroup.add("edittext", undefined, "");
        etPageRange.characters = PAGE_RANGE_CHARS;
        etPageRange.helpTip = getLabel(LABELS.tooltip.pageRange);

        return { etPageRange: etPageRange };
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

        var cropModeLabel = placementPanel.add("statictext", undefined, labelText(LABELS.fieldLabel.cropMode));
        var ddCropMode = placementPanel.add("dropdownlist", undefined, cropNames);
        ddCropMode.minimumSize.width = CROP_LIST_MIN_WIDTH;
        ddCropMode.selection = DEFAULT_CROP_INDEX;
        cropModeLabel.helpTip = ddCropMode.helpTip = getLabel(LABELS.tooltip.cropMode);
        /* トリミング指定はPDFのときだけ使うので、ファイル確定まで無効にする */
        ddCropMode.enabled = false;

        var evenPageGroup = setupRowGroup(placementPanel.add("group"));
        var evenPageLabel = evenPageGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.evenPage));
        var rbEvenPageRight = evenPageGroup.add("radiobutton", undefined, getLabel(LABELS.radio.evenPageRight));
        var rbEvenPageLeft = evenPageGroup.add("radiobutton", undefined, getLabel(LABELS.radio.evenPageLeft));
        rbEvenPageRight.value = true;
        evenPageLabel.helpTip = rbEvenPageRight.helpTip = rbEvenPageLeft.helpTip = getLabel(LABELS.tooltip.evenPage);

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
     * 縦並びの列グループを追加します。
     * @param {Group} parentGroup - 追加先のグループ
     * @returns {Group} 追加した列グループ
     */
    function addColumnGroup(parentGroup) {
        var columnGroup = parentGroup.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = "fill";
        return columnGroup;
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

        var leftColumnGroup = addColumnGroup(columnsGroup);
        var rightColumnGroup = addColumnGroup(columnsGroup);

        var sourceControls = buildSourcePanel(leftColumnGroup);
        var newDocumentControls = buildNewDocumentPanel(rightColumnGroup, sourceDoc);
        var pagesControls = buildPagesPanel(leftColumnGroup);
        var placementControls = buildPlacementPanel(rightColumnGroup);
        buildButtonRow(dialogWindow);

        return {
            dialogWindow: dialogWindow,
            btnSelectFile: sourceControls.btnSelectFile,
            sourceNameText: sourceControls.sourceNameText,
            rbColorRGB: newDocumentControls.rbColorRGB,
            etPageRange: pagesControls.etPageRange,
            ddCropMode: placementControls.ddCropMode,
            rbEvenPageRight: placementControls.rbEvenPageRight,
            rbEvenPageLeft: placementControls.rbEvenPageLeft
        };
    }

    /**
     * ダイアログの設定値から読み込み条件をまとめます。
     * @param {object} dialogControls - buildDialog() が返すコントロールの参照
     * @param {File} sourceFile - 読み込むPDF/AIファイル
     * @returns {ImportSettings} 読み込み条件
     */
    function readImportSettings(dialogControls, sourceFile) {
        var pageNumbers = parsePageNumbers(dialogControls.etPageRange.text);
        if (pageNumbers.length === 0) pageNumbers = [1];

        var cropIndex = dialogControls.ddCropMode.selection ? dialogControls.ddCropMode.selection.index : DEFAULT_CROP_INDEX;

        return {
            sourceFile: sourceFile,
            pageNumbers: pageNumbers,
            cropMode: CROP_OPTIONS[cropIndex].value,
            colorSpace: dialogControls.rbColorRGB.value ? DocumentColorSpace.RGB : DocumentColorSpace.CMYK,
            evenPageOnRight: dialogControls.rbEvenPageRight.value
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
        var dialogControls = buildDialog(sourceDoc);

        /* 現在の読み込み対象ファイル。未選択時は null。 */
        var sourceFile = null;

        /**
         * 読み込み対象ファイルを差し替え、ファイル名・ページ範囲・トリミング指定・綴じ方向を更新します。
         * ページ数が取れないときはアラートを出し、ページ範囲を空にします。
         * @param {File} pdfAiFile - 読み込み対象のPDF/AIファイル
         * @returns {void}
         */
        function applySourceFile(pdfAiFile) {
            var pageCount = readPageCountFromFile(pdfAiFile);
            if (!pageCount) showAlert(LABELS.alert.pageCountFailed);

            sourceFile = pdfAiFile;

            /* トリミング指定はPDFのときのみ有効 */
            dialogControls.ddCropMode.enabled = isPdfFile(sourceFile);

            /* PDFの綴じ方向から偶数ページの位置を自動設定 */
            if (isRightBoundFile(sourceFile)) dialogControls.rbEvenPageRight.value = true;
            else dialogControls.rbEvenPageLeft.value = true;

            dialogControls.sourceNameText.text = decodeSafely(sourceFile.name);
            dialogControls.sourceNameText.helpTip = decodeSafely(sourceFile.fsName);
            dialogControls.etPageRange.text = pageCount ? ("1-" + pageCount) : "";
        }

        dialogControls.btnSelectFile.onClick = function () {
            var pickedFile = File.openDialog(getLabel(LABELS.dialog.pickFile), getLabel(LABELS.dialog.pickFilter));
            if (!pickedFile) return;

            if (isPdfOrAiFile(pickedFile)) {
                applySourceFile(pickedFile);
            } else {
                showAlert(LABELS.alert.pickPdfAi);
                dialogControls.etPageRange.text = "";
            }
        };

        /* 選択中の配置画像がPDF/AIなら、そのリンク先を初期値にする */
        var selectedPlacedItem = findFirstPlacedItem(sourceDoc.selection);
        if (selectedPlacedItem) {
            if (!selectedPlacedItem.file) showAlert(LABELS.alert.linkUnknown);
            else if (isPdfOrAiFile(selectedPlacedItem.file)) applySourceFile(selectedPlacedItem.file);
        }

        if (dialogControls.dialogWindow.show() !== 1) return;

        if (!sourceFile) {
            showAlert(LABELS.alert.needFile);
            return;
        }

        importPages(sourceDoc, readImportSettings(dialogControls, sourceFile));
    }

    main();

})();
