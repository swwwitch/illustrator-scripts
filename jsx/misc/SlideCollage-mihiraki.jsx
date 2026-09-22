#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

PDF/AIファイルを指定したページ範囲で配置し、各ページを個別のアートボードとして展開します。
横長ページは見開きとみなして左右に分割し、綴じ方向に応じた順序で配置できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SlideCollage-mihiraki.md

### Overview

Places a PDF/AI file over a given page range and lays each page out on its own artboard.
Landscape pages are treated as spreads, split left and right, and ordered according to the binding direction.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SlideCollage-mihiraki.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SlideCollage-mihiraki";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SlideCollage-mihiraki.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SlideCollage-mihiraki.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var ARTBOARD_GAP = 20;            /* アートボード間の間隔（pt） / Gap between artboards (pt) */
    var SPREAD_ASPECT_RATIO = 1.2;    /* 幅が高さのこの倍率を超えたら見開きとみなす / Wider than height x this = spread */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS = [15, 20, 15, 10];   /* パネルの余白 / Panel margins */
    var SOURCE_NAME_CHARACTERS = 20;        /* ファイル名欄の幅（文字数） / Width of the file name field */
    var PAGE_RANGE_CHARACTERS = 10;         /* 範囲欄の幅（文字数） / Width of the page range field */
    var CROP_DROPDOWN_MIN_WIDTH = 160;      /* トリミングのドロップダウンの最小幅 / Minimum width of the crop dropdown */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "スライドコラージュ（見開き対応）", en: "Slide Collage (Spread Support)" }
        },
        panel: {
            artboards: { ja: "アートボード", en: "Artboards" },
            sourceFile: { ja: "読み込みファイル", en: "Source File" },
            items: { ja: "アイテム", en: "Items" }
        },
        fieldLabel: {
            range: { ja: "指定", en: "Range" }
        },
        radio: {
            bindR2L: { ja: "右綴じ", en: "Right-to-Left" },
            bindL2R: { ja: "左綴じ", en: "Left-to-Right" }
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
        fileDialog: {
            prompt: { ja: "PDF/AIを選択してください", en: "Select a PDF/AI" },
            filter: { ja: "PDF/AI:*.pdf;*.ai", en: "PDF/AI:*.pdf;*.ai" }
        },
        status: {
            notSelected: { ja: "未指定", en: "Not selected" }
        },
        tooltip: {
            selectFile: {
                ja: "PDF または AI ファイルを選びます。ページ数を調べて、範囲に「1-最終ページ」を入れます",
                en: "Choose a PDF or AI file; the range is filled with 1 to its last page"
            },
            pageRange: {
                ja: "配置するページを指定します（例：1-5, 8, 10-12）。空欄のときは1ページ目だけを配置します",
                en: "Pages to place (e.g. 1-5, 8, 10-12). Leave empty to place page 1 only"
            },
            cropBox: {
                ja: "PDF を配置するときの範囲を選びます。AI ファイルでは使えません",
                en: "Choose the PDF box used when placing pages. Not available for AI files"
            },
            bindR2L: {
                ja: "見開き（横長）ページを右半分、左半分の順に並べます。PDF に右綴じの指定があれば、ファイルを選んだときに自動で選ばれます",
                en: "Places the right half of each spread (landscape page) first. Chosen automatically when the PDF specifies right-to-left binding"
            },
            bindL2R: {
                ja: "見開き（横長）ページを左半分、右半分の順に並べます",
                en: "Places the left half of each spread (landscape page) first"
            }
        },
        alert: {
            linkUnknown: { ja: "画像のリンク先が不明でした。", en: "Image link not found." },
            pageCountFail: {
                ja: "リンクされたPDF/AIファイルのページ数を取得できませんでした。",
                en: "Could not determine the page count of the linked PDF/AI file."
            },
            pickPdfAi: { ja: "PDFまたはAIファイルを選択してください。", en: "Please select a PDF or AI file." },
            needDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document before running." },
            placeError: { ja: "配置中にエラーが発生しました。", en: "An error occurred while placing the pages." },
            needFile: {
                ja: "先に［ファイル指定］で読み込みファイルを選択してください。",
                en: "Please select a source file first."
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) {
                return labelPath;
            }
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // PDF の読み込み設定 / PDF import settings
    // =========================================

    /* トリミングの種類ごとの環境設定値 / Preference value for each crop box */
    var CROP_MODE = {
        art: 4,
        trim: 3,
        bleed: 2,
        crop: 1
    };

    /**
     * PDF 読み込み時のトリミング（クロップボックス）の環境設定を書き込む
     * Illustrator のバージョン差を吸収するため、候補のキーすべてに書き込む
     * 期待する値（多くの環境で）：0=Media, 1=Crop, 2=Bleed, 3=Trim, 4=Art
     * @param {number} cropMode - CROP_MODE の値
     * @returns {void}
     */
    function setPdfCropPreference(cropMode) {
        var cropPrefKeys = [
            "plugin/PDFImport/CropToBox",
            "plugin/PDFImport/CropTo",
            "plugin/PDFImport/CropBox",
            "plugin/PDFImport/CropToType"
        ];
        for (var i = 0; i < cropPrefKeys.length; i++) {
            /* 環境によって存在しないキーがある / Some keys do not exist in every version */
            try {
                app.preferences.setIntegerPreference(cropPrefKeys[i], cropMode);
            } catch (e) { }
        }
    }

    /**
     * 配置するページ番号を環境設定に書き込む（AI も PDF の読み込みで配置されるので同じキーを使う）
     * @param {number|string} pageNumber - ページ番号（1未満や数値でないときは1）
     * @returns {void}
     */
    function setImportPageNumber(pageNumber) {
        var validPageNumber = parseInt(pageNumber, 10);
        if (isNaN(validPageNumber) || validPageNumber < 1) validPageNumber = 1;
        try {
            app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", validPageNumber);
        } catch (e) { }
    }

    /**
     * 次の配置に使うトリミングとページ番号を環境設定に書き込む
     * @param {File} sourceFile - 読み込みファイル
     * @param {number} pageNumber - ページ番号
     * @param {number} cropMode - CROP_MODE の値
     * @returns {void}
     */
    function prepareImport(sourceFile, pageNumber, cropMode) {
        if (isPdfLikeFile(sourceFile)) {
            setPdfCropPreference(cropMode);
        }
        setImportPageNumber(pageNumber);
    }

    // =========================================
    // ファイルの判定 / File inspection
    // =========================================

    /**
     * PDF 互換として扱うファイルか（AI も配置時は PDF の読み込みを通るので含める）
     * @param {File} file - 判定するファイル
     * @returns {boolean} 名前に .pdf / .ai を含めば true
     */
    function isPdfLikeFile(file) {
        if (!file) return false;
        var fileName = (file.name || "").toLowerCase();
        return (fileName.indexOf(".pdf") > -1) || (fileName.indexOf(".ai") > -1);
    }

    /**
     * PDF ファイルか（トリミングの指定は PDF でだけ意味があるので UI の有効化に使う）
     * @param {File} file - 判定するファイル
     * @returns {boolean} 名前に .pdf を含めば true
     */
    function isPdfFile(file) {
        if (!file) return false;
        return ((file.name || "").toLowerCase().indexOf(".pdf") > -1);
    }

    /**
     * PDF の綴じ方向を判定する（先頭 200 行に /Direction /R2L があれば右綴じ）
     * @param {File} file - PDF ファイル
     * @returns {string} "R2L"（右綴じ）または "L2R"（左綴じ）
     */
    function detectPdfBindingDirection(file) {
        try {
            file.open("r");
            var lineCount = 0;
            while (!file.eof && lineCount < 200) {
                var line = file.readln();
                lineCount++;
                if (/\/Direction\s*\/R2L/.test(line)) {
                    file.close();
                    return "R2L";
                }
            }
            file.close();
        } catch (e) {
            try { file.close(); } catch (e2) { }
        }
        return "L2R";
    }

    /**
     * PDF / AI ファイルの中身から総ページ数を推定する
     * @param {File} file - PDF / AI ファイル
     * @returns {number|undefined} 推定したページ数（読み取れなければ undefined）
     */
    function estimatePageCountFromFile(file) {
        var countPattern = /<<\/Count\s(\d+)/;
        var pageParentPattern = /<<\/Type\/Page\/Parent/;
        var pageTypePattern = /\/Type\s\/Page\s/;
        var structPagePattern = /\/StructParents\s\d+.*\/Type\/Page>>/;
        var linearizedPattern = /<<\/Linearized\s.+\/N\s(\d+)\/T\s.+>>/;
        var pagesTypePattern = /\/Type\/Pages/;
        var countValuePattern = /\/Count\s(\d+)/;

        var pageCount;
        var pageObjectCount = 0;
        try {
            file.open("r");
            while (!file.eof) {
                var line = file.readln();
                if (countPattern.test(line) || linearizedPattern.test(line)) { pageCount = Number(RegExp.$1); break; }
                if (pageParentPattern.test(line) || pageTypePattern.test(line) || structPagePattern.test(line)) pageObjectCount++;
                if (pagesTypePattern.test(line)) {
                    line = file.readln();
                    if (countValuePattern.test(line)) {
                        var pagesCount = Number(RegExp.$1);
                        if (pageObjectCount < pagesCount) pageObjectCount = pagesCount;
                    }
                }
            }
            if (pageObjectCount > 0) pageCount = pageObjectCount;
        } catch (e) {
            alert(e);
        } finally {
            try { file.close(); } catch (e2) { }
        }
        return pageCount;
    }

    /**
     * 配置画像の中から最初の PlacedItem を探す（グループの中もたどる）
     * @param {PageItem[]} items - 探す対象
     * @returns {PlacedItem|null} 見つかった配置画像
     */
    function findFirstPlacedItem(items) {
        if (!items || items.length <= 0) return null;
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) continue;
            var constructorName = (item.constructor && item.constructor.name) ? item.constructor.name : "";
            if (constructorName === "PlacedItem") return item;
            if (constructorName === "GroupItem") {
                var nestedPlacedItem = findFirstPlacedItem(item.pageItems);
                if (nestedPlacedItem) return nestedPlacedItem;
            }
        }
        return null;
    }

    /**
     * オブジェクトの中の配置画像（PDF / AI）の最終ページ番号を返す（再リンクや配置はしない）
     * @param {PageItem[]} items - 探す対象
     * @returns {number|null} 最終ページ番号（取得できなければ null）
     */
    function getLastPageOfPlacedFile(items) {
        var placedItem = findFirstPlacedItem(items);
        if (!placedItem) return null;

        var linkedFile = placedItem.file;
        if (!linkedFile) {
            alert(getLabel("alert.linkUnknown"));
            return null;
        }

        if (!/\.(?:pdf|ai)$/i.test(decodeURIComponent(linkedFile.name))) return null;

        var lastPage = estimatePageCountFromFile(linkedFile);
        if (!lastPage || isNaN(Number(lastPage)) || Number(lastPage) <= 0) {
            alert(getLabel("alert.pageCountFail"));
            return null;
        }
        return Number(lastPage);
    }

    /**
     * 選んだファイル、または選択中の配置画像から読み込みファイルと最終ページを取得し、コールバックで返す
     * ファイルが指定されたときは一時的に配置して、選択からの取得と同じ処理で調べる
     * @param {Document} doc - 対象ドキュメント
     * @param {File|null} pickedFile - ユーザーが選んだファイル（null なら選択中の配置画像を使う）
     * @param {Function} onSourceFile - 読み込みファイルを受け取るコールバック
     * @param {Function} onLastPage - 最終ページ番号（または null）を受け取るコールバック
     * @returns {void}
     */
    function loadSourceFromFileOrSelection(doc, pickedFile, onSourceFile, onLastPage) {
        try {
            if (pickedFile) {
                if (!/\.(?:pdf|ai)$/i.test(decodeURIComponent(pickedFile.name))) {
                    alert(getLabel("alert.pickPdfAi"));
                    onLastPage(null);
                    return;
                }

                var tempPlacedItem = null;
                try {
                    tempPlacedItem = doc.placedItems.add();
                    tempPlacedItem.file = pickedFile;

                    /* 表示中の範囲の左上に置く（ビューが取れないことがある） / Put it at the top-left of the view (the view may be unavailable) */
                    try {
                        var viewBounds = doc.activeView && doc.activeView.bounds ? doc.activeView.bounds : null; /* [left, top, right, bottom] */
                        if (viewBounds && viewBounds.length === 4) {
                            tempPlacedItem.position = [viewBounds[0], viewBounds[1]];
                        }
                    } catch (e) { }

                    var lastPageOfPicked = getLastPageOfPlacedFile([tempPlacedItem]);
                    onSourceFile(pickedFile);
                    onLastPage(lastPageOfPicked);
                } finally {
                    if (tempPlacedItem) {
                        try { tempPlacedItem.remove(); } catch (e2) { }
                    }
                }
                return;
            }

            /* ファイル未指定：選択中の配置画像を使う / No file given: use the selected placed item */
            var lastPage = getLastPageOfPlacedFile(doc.selection);

            /* リンク切れの画像は .file で例外になることがある / .file can throw for a missing link */
            try {
                var selectedPlacedItem = findFirstPlacedItem(doc.selection);
                if (selectedPlacedItem && selectedPlacedItem.file) onSourceFile(selectedPlacedItem.file);
            } catch (e3) { }

            onLastPage(lastPage);
        } catch (e4) {
            alert(e4);
            onLastPage(null);
        }
    }

    /**
     * 「1-20」「1,3,5」のような文字列をページ番号の配列にする
     * @param {string} rangeText - ページ範囲の文字列
     * @returns {number[]} ページ番号の配列（範囲は小さい方から展開）
     */
    function parsePageNumbers(rangeText) {
        var pageNumbers = [];
        var rangeParts = rangeText.split(",");

        for (var i = 0; i < rangeParts.length; i++) {
            var rangePart = rangeParts[i].replace(/^\s+|\s+$/g, "");

            if (rangePart.indexOf("-") > -1) {
                var rangeEnds = rangePart.split("-");
                var rangeStart = parseInt(rangeEnds[0], 10);
                var rangeEnd = parseInt(rangeEnds[1], 10);

                if (!isNaN(rangeStart) && !isNaN(rangeEnd)) {
                    var lastPage = Math.max(rangeStart, rangeEnd);
                    for (var j = Math.min(rangeStart, rangeEnd); j <= lastPage; j++) {
                        pageNumbers.push(j);
                    }
                }
            } else {
                var pageNumber = parseInt(rangePart, 10);
                if (!isNaN(pageNumber)) {
                    pageNumbers.push(pageNumber);
                }
            }
        }
        return pageNumbers;
    }

    // =========================================
    // ダイアログの位置（セッション中のみ） / Dialog position (session only)
    // =========================================

    /* $.global に置くキー（変えると記憶した位置が読めなくなる） / $.global key; do not rename */
    var DIALOG_BOUNDS_KEY = "__SC_slideCollage_dialog_bounds";

    /**
     * 有限の数値か
     * @param {*} value - 調べる値
     * @returns {boolean} 有限の数値なら true
     */
    function isFiniteNumber(value) {
        return (typeof value === "number") && isFinite(value);
    }

    /**
     * 位置を素のオブジェクトに写す（4辺とも有限の数値でなければ null）
     * @param {Object|null} bounds - left / top / right / bottom を持つ値
     * @returns {Object|null} { left, top, right, bottom } または null
     */
    function toFiniteBounds(bounds) {
        if (!bounds) return null;
        var plainBounds = { left: bounds.left, top: bounds.top, right: bounds.right, bottom: bounds.bottom };
        if (!isFiniteNumber(plainBounds.left) || !isFiniteNumber(plainBounds.top) ||
            !isFiniteNumber(plainBounds.right) || !isFiniteNumber(plainBounds.bottom)) return null;
        return plainBounds;
    }

    /**
     * 記憶しているダイアログの位置を返す
     * @returns {Object|undefined} 記憶している位置
     */
    function loadDialogBounds() {
        return $.global[DIALOG_BOUNDS_KEY];
    }

    /**
     * ダイアログの位置を記憶する（ScriptUI のオブジェクトではなく素のオブジェクトで保存）
     * @param {Bounds} bounds - ダイアログの bounds
     * @returns {void}
     */
    function saveDialogBounds(bounds) {
        var plainBounds = toFiniteBounds(bounds);
        if (plainBounds) $.global[DIALOG_BOUNDS_KEY] = plainBounds;
    }

    /**
     * 位置がどれかの画面と 40px 以上重なっているか
     * @param {Object} bounds - { left, top, right, bottom }
     * @returns {boolean} 重なっていれば true（画面の情報が取れないときも true）
     */
    function isOnAnyScreen(bounds) {
        /* 画面の情報が取れないときはそのまま適用する / Apply as is when screen info is unavailable */
        try {
            if (!($.screens && $.screens.length)) return false;
            for (var i = 0; i < $.screens.length; i++) {
                var screenBounds = $.screens[i].bounds; /* [L, T, R, B] */
                var overlapWidth = Math.min(bounds.right, screenBounds[2]) - Math.max(bounds.left, screenBounds[0]);
                var overlapHeight = Math.min(bounds.bottom, screenBounds[3]) - Math.max(bounds.top, screenBounds[1]);
                if (overlapWidth > 40 && overlapHeight > 40) return true;
            }
            return false;
        } catch (e) {
            return true;
        }
    }

    /**
     * 記憶している位置をダイアログに適用する（壊れた値や画面外の位置は無視）
     * @param {Window} targetDialog - 対象のダイアログ
     * @returns {void}
     */
    function applyDialogBounds(targetDialog) {
        var savedBounds = toFiniteBounds(loadDialogBounds());
        if (!savedBounds) return;

        /* 小さすぎる値は壊れているとみなす / Too small means invalid */
        var savedWidth = savedBounds.right - savedBounds.left;
        var savedHeight = savedBounds.bottom - savedBounds.top;
        if (!(savedWidth >= 200 && savedHeight >= 120)) return;

        if (!isOnAnyScreen(savedBounds)) return;

        /* 適用できなければ既定の位置のまま / Keep the default position if it cannot be applied */
        try {
            targetDialog.bounds = [savedBounds.left, savedBounds.top, savedBounds.right, savedBounds.bottom];
        } catch (e) { }
    }

    /**
     * セッション中のダイアログ位置を復元し、動かしたら記憶する
     * @param {Window} targetDialog - 対象のダイアログ
     * @returns {void}
     */
    function trackDialogPosition(targetDialog) {
        applyDialogBounds(targetDialog);
        /* 記憶が無いときは中央に / Center when nothing is stored */
        if (!loadDialogBounds()) {
            targetDialog.center();
        }
        targetDialog.onMove = function () { saveDialogBounds(targetDialog.bounds); };
        targetDialog.onMoving = function () { saveDialogBounds(targetDialog.bounds); };
    }

    // =========================================
    // 配置 / Placement
    // =========================================

    /**
     * アートボードの範囲でクリッピングマスクをかける
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} item - マスクをかけるオブジェクト
     * @param {number[]} artboardRect - アートボードの範囲 [left, top, right, bottom]
     * @returns {GroupItem} クリップグループ
     */
    function clipToArtboard(doc, item, artboardRect) {
        var maskWidth = Math.abs(artboardRect[2] - artboardRect[0]);
        var maskHeight = Math.abs(artboardRect[1] - artboardRect[3]);
        if (maskWidth <= 0) maskWidth = 1;
        if (maskHeight <= 0) maskHeight = 1;

        var maskPath = doc.activeLayer.pathItems.rectangle(artboardRect[1], artboardRect[0], maskWidth, maskHeight);
        maskPath.stroked = false;
        maskPath.filled = false;
        maskPath.clipping = true;

        var clipGroup = doc.groupItems.add();
        try { item.moveToEnd(clipGroup); } catch (e) { }
        try { maskPath.moveToBeginning(clipGroup); } catch (e2) { }

        clipGroup.clipped = true;
        return clipGroup;
    }

    /**
     * ページを一時的に配置して寸法を測る（測ったら削除）
     * @param {Document} doc - 対象ドキュメント
     * @param {File} sourceFile - 読み込みファイル
     * @param {number} pageNumber - ページ番号
     * @param {number} cropMode - CROP_MODE の値
     * @returns {{width: number, height: number}} ページの寸法
     */
    function measurePlacedPageSize(doc, sourceFile, pageNumber, cropMode) {
        prepareImport(sourceFile, pageNumber, cropMode);

        var measureItem = null;
        try {
            measureItem = doc.placedItems.add();
            measureItem.file = sourceFile;
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
     * アートボードを1枚用意し（最初は作業中のアートボードを使い、以降は追加）、ページを配置してマスクをかける
     * @param {Object} placement - 配置の状態（doc / sourceFile / cropMode / activeIndex / top / artboardCount）
     * @param {number} pageNumber - ページ番号
     * @param {number[]} artboardRect - アートボードの範囲 [left, top, right, bottom]
     * @param {number} positionX - 配置画像の左端の X
     * @returns {void}
     */
    function placeOnArtboard(placement, pageNumber, artboardRect, positionX) {
        var doc = placement.doc;
        if (placement.artboardCount === 0) {
            doc.artboards[placement.activeIndex].artboardRect = artboardRect;
        } else {
            doc.artboards.add(artboardRect);
        }
        placement.artboardCount++;

        prepareImport(placement.sourceFile, pageNumber, placement.cropMode);
        var placedPage = doc.placedItems.add();
        placedPage.file = placement.sourceFile;
        placedPage.position = [positionX, placement.top];
        clipToArtboard(doc, placedPage, artboardRect);
    }

    /**
     * 単ページを1枚のアートボードに配置する
     * @param {Object} placement - 配置の状態（nextX を更新する）
     * @param {number} pageNumber - ページ番号
     * @param {{width: number, height: number}} pageSize - ページの寸法
     * @returns {void}
     */
    function placeSinglePage(placement, pageNumber, pageSize) {
        var pageX = placement.nextX;
        var pageRect = [pageX, placement.top, pageX + pageSize.width, placement.top - pageSize.height];
        placeOnArtboard(placement, pageNumber, pageRect, pageX);
        placement.nextX = pageX + pageSize.width + ARTBOARD_GAP;
    }

    /**
     * 見開きページを左右に分けて2枚のアートボードに配置する
     * 右綴じでは右半分、左綴じでは左半分を先（左側のアートボード）に置く
     * @param {Object} placement - 配置の状態（nextX を更新する）
     * @param {number} pageNumber - ページ番号
     * @param {{width: number, height: number}} pageSize - 見開きの寸法
     * @param {boolean} isRightBinding - 右綴じなら true
     * @returns {void}
     */
    function placeSpreadPage(placement, pageNumber, pageSize, isRightBinding) {
        var halfWidth = pageSize.width / 2;
        var pageBottom = placement.top - pageSize.height;

        var firstX = placement.nextX;
        var firstRect = [firstX, placement.top, firstX + halfWidth, pageBottom];
        placeOnArtboard(placement, pageNumber, firstRect, isRightBinding ? (firstX - halfWidth) : firstX);

        var secondX = firstX + halfWidth + ARTBOARD_GAP;
        var secondRect = [secondX, placement.top, secondX + halfWidth, pageBottom];
        placeOnArtboard(placement, pageNumber, secondRect, isRightBinding ? secondX : (secondX - halfWidth));

        placement.nextX = secondX + halfWidth + ARTBOARD_GAP;
    }

    /**
     * 各ページを個別のアートボードに配置する。見開き（横長）ページは左右に分けて2つのアートボードに配置
     * @param {Document} doc - 対象ドキュメント
     * @param {File} sourceFile - 読み込みファイル
     * @param {number[]} targetPages - 配置するページ番号
     * @param {number} cropMode - CROP_MODE の値
     * @param {boolean} isRightBinding - 右綴じなら true
     * @returns {void}
     */
    function placeOnIndividualArtboards(doc, sourceFile, targetPages, cropMode, isRightBinding) {
        var activeIndex = doc.artboards.getActiveArtboardIndex();
        var baseRect = doc.artboards[activeIndex].artboardRect;
        var placement = {
            doc: doc,
            sourceFile: sourceFile,
            cropMode: cropMode,
            activeIndex: activeIndex,
            top: baseRect[1],
            nextX: baseRect[0],
            artboardCount: 0    /* 作成・使用したアートボード数 / Artboards created or reused */
        };

        try {
            for (var i = 0; i < targetPages.length; i++) {
                var pageNumber = parseInt(targetPages[i], 10);
                if (isNaN(pageNumber) || pageNumber < 1) pageNumber = 1;

                var pageSize = measurePlacedPageSize(doc, sourceFile, pageNumber, cropMode);
                if (pageSize.width > pageSize.height * SPREAD_ASPECT_RATIO) {
                    placeSpreadPage(placement, pageNumber, pageSize, isRightBinding);
                } else {
                    placeSinglePage(placement, pageNumber, pageSize);
                }
            }
        } catch (e) {
            alert(getLabel("alert.placeError"));
        } finally {
            setImportPageNumber(1);
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 項目を縦に並べるパネルを追加する
     * @param {Group} parentGroup - 追加先
     * @param {string} titlePath - パネルタイトルの LABELS パス
     * @returns {Panel} 追加したパネル
     */
    function addColumnPanel(parentGroup, titlePath) {
        var columnPanel = parentGroup.add("panel", undefined, getLabel(titlePath));
        columnPanel.orientation = "column";
        columnPanel.alignChildren = ["left", "top"];
        columnPanel.margins = PANEL_MARGINS;
        return columnPanel;
    }

    /**
     * アイテムパネル（PDF のトリミングと綴じ方向）を追加する
     * @param {Group} parentGroup - 追加先
     * @returns {{cropDropdown: DropDownList, rightBindingRadio: RadioButton, leftBindingRadio: RadioButton}} 追加したコントロール
     */
    function addItemPanel(parentGroup) {
        var itemPanel = parentGroup.add("panel", undefined, getLabel("panel.items"));
        itemPanel.alignChildren = "left";
        itemPanel.margins = PANEL_MARGINS;

        var cropDropdown = itemPanel.add("dropdownlist", undefined, [
            getLabel("dropdown.cropArt"), getLabel("dropdown.cropTrim"), getLabel("dropdown.cropCrop"), getLabel("dropdown.cropBleed")
        ]);
        cropDropdown.minimumSize.width = CROP_DROPDOWN_MIN_WIDTH;
        cropDropdown.helpTip = getLabel("tooltip.cropBox");
        /* 既定：仕上がり / Default: Crop */
        cropDropdown.selection = 2;

        /* 綴じ方向（右綴じ / 左綴じ） / Binding direction */
        var bindingGroup = itemPanel.add("group");
        bindingGroup.orientation = "row";
        bindingGroup.alignChildren = ["left", "center"];

        var rightBindingRadio = bindingGroup.add("radiobutton", undefined, getLabel("radio.bindR2L"));
        var leftBindingRadio = bindingGroup.add("radiobutton", undefined, getLabel("radio.bindL2R"));
        rightBindingRadio.helpTip = getLabel("tooltip.bindR2L");
        leftBindingRadio.helpTip = getLabel("tooltip.bindL2R");
        /* 既定：右綴じ / Default: right-to-left */
        rightBindingRadio.value = true;

        return {
            cropDropdown: cropDropdown,
            rightBindingRadio: rightBindingRadio,
            leftBindingRadio: leftBindingRadio
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを表示し、OK なら指定ページを個別のアートボードに配置する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.needDocument"));
            return;
        }

        var doc = app.activeDocument;
        /* ユーザーが選んだ読み込みファイル（未指定なら null） / Source file chosen by the user (null until chosen) */
        var sourceFile = null;

        var collageDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        collageDialog.alignChildren = "fill";
        trackDialogPosition(collageDialog);

        /* 単カラム構成 / Single column */
        var mainColumn = collageDialog.add("group");
        mainColumn.orientation = "column";
        mainColumn.alignChildren = "fill";

        /* 読み込みファイル / Source file */
        var sourcePanel = addColumnPanel(mainColumn, "panel.sourceFile");
        var btnSelectFile = sourcePanel.add("button", undefined, getLabel("button.selectFile"));
        btnSelectFile.helpTip = getLabel("tooltip.selectFile");
        var sourceNameText = sourcePanel.add("statictext", undefined, getLabel("status.notSelected"));
        sourceNameText.characters = SOURCE_NAME_CHARACTERS;

        /* アートボード（ページ範囲） / Artboards (page range) */
        var artboardPanel = addColumnPanel(mainColumn, "panel.artboards");
        var rangeRowGroup = artboardPanel.add("group");
        rangeRowGroup.orientation = "row";
        rangeRowGroup.alignChildren = ["left", "center"];
        rangeRowGroup.add("statictext", undefined, labelText("fieldLabel.range"));
        var pageRangeInput = rangeRowGroup.add("edittext", undefined, "");
        pageRangeInput.characters = PAGE_RANGE_CHARACTERS;
        pageRangeInput.helpTip = getLabel("tooltip.pageRange");

        var itemControls;

        /**
         * 読み込みファイルを差し替え、トリミングの有効／無効・綴じ方向・ファイル名の表示を更新する
         * @param {File|null} file - 読み込みファイル
         * @returns {void}
         */
        function setSourceFile(file) {
            sourceFile = file || null;

            /* トリミングは PDF でだけ意味がある / Crop box only matters for PDF */
            itemControls.cropDropdown.enabled = isPdfFile(sourceFile);

            applyDetectedBinding(file);

            /* 名前に % が残っていると decodeURIComponent が例外 / decodeURIComponent throws on a stray % */
            try {
                if (file) {
                    sourceNameText.text = decodeURIComponent(file.name);
                    sourceNameText.helpTip = decodeURIComponent(file.fsName);
                } else {
                    sourceNameText.text = getLabel("status.notSelected");
                    sourceNameText.helpTip = "";
                }
            } catch (e) {
                sourceNameText.text = file ? String(file.name) : getLabel("status.notSelected");
                sourceNameText.helpTip = file ? String(file.fsName) : "";
            }
        }

        /**
         * 範囲欄に「1-最終ページ」を入れる（取得できなければ空欄）
         * @param {number|null} lastPage - 最終ページ番号
         * @returns {void}
         */
        function setPageRangeText(lastPage) {
            pageRangeInput.text = lastPage ? ("1-" + lastPage) : "";
        }

        /**
         * ファイルの綴じ方向を判定してラジオボタンに反映する
         * @param {File|null} file - 読み込みファイル
         * @returns {void}
         */
        function applyDetectedBinding(file) {
            if (!file) return;
            if (detectPdfBindingDirection(file) === "R2L") {
                itemControls.rightBindingRadio.value = true;
            } else {
                itemControls.leftBindingRadio.value = true;
            }
        }

        /**
         * 右綴じが選ばれているか
         * @returns {boolean} 右綴じなら true
         */
        function isRightBinding() {
            return !!itemControls.rightBindingRadio.value;
        }

        /**
         * ドロップダウンで選んだトリミングを CROP_MODE の値で返す
         * @returns {number} CROP_MODE の値
         */
        function getSelectedCropMode() {
            /* 0:アート / 1:トリミング / 2:仕上がり / 3:裁ち落とし / 0: Art, 1: Trim, 2: Crop, 3: Bleed */
            var cropIndex = (itemControls.cropDropdown.selection) ? itemControls.cropDropdown.selection.index : 2;
            if (cropIndex === 0) return CROP_MODE.art;
            if (cropIndex === 1) return CROP_MODE.trim;
            if (cropIndex === 3) return CROP_MODE.bleed;
            /* 仕上がりは CropBox を想定（環境差で失敗しても無視される） / Crop maps to CropBox; failures are ignored */
            return CROP_MODE.crop;
        }

        /* 初期表示：選択中の配置画像から読み込む / Initial state: read from the selected placed item */
        loadSourceFromFileOrSelection(doc, null, setSourceFile, setPageRangeText);
        itemControls.cropDropdown.enabled = isPdfFile(sourceFile);

        btnSelectFile.onClick = function () {
            var pickedFile = File.openDialog(getLabel("fileDialog.prompt"), getLabel("fileDialog.filter"));
            if (!pickedFile) return;
            loadSourceFromFileOrSelection(doc, pickedFile, setSourceFile, setPageRangeText);
        };

        /* アイテム（PDF のトリミングと綴じ方向） / Items (PDF crop box and binding) */
        itemControls = addItemPanel(mainColumn);
        itemControls.cropDropdown.enabled = isPdfFile(sourceFile);

        /* ボタン（キャンセル・OK） / Buttons (Cancel, OK) */
        var btnRowGroup = collageDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["center", "center"];

        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /**
         * 位置を記憶してからダイアログを閉じる
         * @param {number} resultCode - show() の戻り値にする値
         * @returns {void}
         */
        function closeDialog(resultCode) {
            saveDialogBounds(collageDialog.bounds);
            collageDialog.close(resultCode);
        }

        btnCancel.onClick = function () {
            closeDialog(2);
        };

        btnOK.onClick = function () {
            closeDialog(1);
        };

        if (collageDialog.show() === 1) {
            if (!sourceFile) {
                alert(getLabel("alert.needFile"));
            } else {
                var targetPages = parsePageNumbers(pageRangeInput.text);
                if (!targetPages || targetPages.length === 0) targetPages = [1];

                placeOnIndividualArtboards(doc, sourceFile, targetPages, getSelectedCropMode(), isRightBinding());
            }
        }

        /* ダイアログ終了後に選択解除 / Deselect after the dialog closes */
        try { doc.selection = null; } catch (e) { }
    }

    main();

})();
