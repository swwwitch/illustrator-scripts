#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * PDFAISpreadImporter.jsx
 *
 * 概要:
 * PDF/AIファイルを指定したページ範囲で読み込み、新規ドキュメント上に各ページを
 * 個別のアートボードとして配置します。
 *
 * 横長ページは見開きとして自動判定し、左右2つのアートボードに分割して配置します。
 * 偶数ページの位置は 右 / 左 から選択でき、見開きページの配置順序に反映されます。
 *
 * PDFではトリミング設定（アート / トリミング / 仕上がり / 裁ち落とし）を選択可能です。
 * 選択中の配置画像、または指定ファイルからページ数を推定し、ページ範囲に反映します。
 *
 * 新規ドキュメントのカラーモードは CMYK / RGB から選択できます。
 * アートボード間隔は 100 pt 固定、ラスタライズ効果解像度は 300 ppi 固定です。
 *
 * 配置完了後は、生成されたすべてのアートボードが見えるように表示倍率を自動調整します。
 *
 * 更新日: 2026-03-18
 */

// =========================================
// バージョンとローカライズ
// =========================================

var SCRIPT_VERSION = "v1.1.0";

// 生成されるアートボード間のデフォルト間隔（pt）
var DEFAULT_ARTBOARD_GAP = 100;

// 新規出力ドキュメントのデフォルト・ラスタライズ効果解像度（ppi）
var DEFAULT_RASTER_EFFECTS_RESOLUTION = 300;


function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 */

var LABELS = {
    dialogTitle: {
        ja: "PDF/AI見開き配置",
        en: "PDF/AI Spread Placement"
    },

    // パネル
    panelLoad: { ja: "アートボード", en: "Artboards" },
    panelSource: { ja: "読み込みファイル", en: "Source File" },
    panelItem: { ja: "配置方法", en: "Placement" },
    panelNewDoc: { ja: "新規ドキュメント", en: "New Document" },

    // 読み込みパネル
    btnLoad: { ja: "ファイル指定", en: "Select File" },
    range: { ja: "指定", en: "Range" },
    dlgPickFile: { ja: "PDF/AIを選択してください", en: "Select a PDF/AI" },
    filterPick: { ja: "PDF/AI:*.pdf;*.ai", en: "PDF/AI:*.pdf;*.ai" },
    alertLinkUnknown: { ja: "画像のリンク先が不明でした。", en: "Image link not found." },
    alertPageCountFail: { ja: "リンクされたPDF/AIファイルのページ数を取得できませんでした。", en: "Could not determine the page count of the linked PDF/AI file." },
    alertPickPdfAi: { ja: "PDFまたはAIファイルを選択してください。", en: "Please select a PDF or AI file." },
    notSelected: { ja: "未指定", en: "Not selected" },

    // 綴じ方向
    bindR2L: { ja: "右", en: "Right" },
    bindL2R: { ja: "左", en: "Left" },
    evenPage: { ja: "偶数ページ:", en: "Even Pages:" },
    colorMode: { ja: "カラーモード", en: "Color Mode" },
    colorModeCMYK: { ja: "CMYK", en: "CMYK" },
    colorModeRGB: { ja: "RGB", en: "RGB" },

    // 配置方法パネル
    cropArt: { ja: "アート", en: "Art" },
    cropTrim: { ja: "トリミング", en: "Trim" },
    cropCrop: { ja: "仕上がり", en: "Crop" },
    cropBleed: { ja: "裁ち落とし", en: "Bleed" },

    // ボタン
    cancel: { ja: "キャンセル", en: "Cancel" },
    ok: { ja: "OK", en: "OK" },

    // ファイル / アラート
    alertNeedDoc: { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document before running." },
    alertPlaceError: { ja: "配置中にエラーが発生しました。", en: "An error occurred while placing the pages." },
    alertNeedFile: { ja: "先に［ファイル指定］で読み込みファイルを選択してください。", en: "Please select a source file first." },
    errorDetails: { ja: "詳細:", en: "Details:" },
};

function L(key) {
    var o = LABELS[key];
    if (!o) return key;
    return o[lang] || o.en || o.ja || key;
}

// アラート補助
function SC_getErrorDetailText(e) {
    if (e === undefined || e === null) return "";
    try {
        if (typeof e === "string") return e;
        if (e && e.message) return String(e.message);
        return String(e);
    } catch (_) {
        return "";
    }
}

function SC_alert(key, e) {
    try {
        var msg = L(key);
        var detail = SC_getErrorDetailText(e);
        if (detail) msg += "\n\n" + L("errorDetails") + "\n" + detail;
        alert(msg);
    } catch (_) { }
}

// ========================
// 配置処理ヘルパー
// - 配置前のページ番号指定
// - 配置サイズの計測
// - アートボード内でのクリッピングマスク適用
// - 単ページ / 見開きページの配置
// - 全アートボードが見える表示倍率への調整
// ========================

function placementSetImportPageNumber(pageNum) {
    var n = parseInt(pageNum, 10);
    if (isNaN(n) || n < 1) n = 1;
    try {
        app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", n);
    } catch (_) { }
}

function placementResetImportPageNumber() {
    try {
        app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", 1);
    } catch (_) { }
}

function placementApplyMaskToArtboard(doc, item, abRect, marginPt) {
    if (!item) return null;
    var m = (typeof marginPt === "number" && !isNaN(marginPt) && marginPt >= 0) ? marginPt : 0;

    var left = abRect[0] + m;
    var top = abRect[1] - m;
    var w = Math.abs(abRect[2] - abRect[0]) - (m * 2);
    var h = Math.abs(abRect[1] - abRect[3]) - (m * 2);
    if (w <= 0) w = 1;
    if (h <= 0) h = 1;

    var maskPath = doc.activeLayer.pathItems.rectangle(top, left, w, h);
    maskPath.stroked = false;
    maskPath.filled = false;
    maskPath.clipping = true;

    var grp = doc.groupItems.add();
    try { item.moveToEnd(grp); } catch (e1) { }
    try { maskPath.moveToBeginning(grp); } catch (e0) { }

    grp.clipped = true;
    return grp;
}

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
            try { measureItem.remove(); } catch (_) { }
        }
    }
}

function placementUseOrAddArtboard(doc, activeIdx, abRect, abCount) {
    if (abCount === 0) {
        doc.artboards[activeIdx].artboardRect = abRect;
    } else {
        doc.artboards.add(abRect);
    }
    return abCount + 1;
}

function placementPlacePageWithMask(doc, fileObj, pageNum, abRect, pos, cropMode) {
    if (isPdfLikeFile(fileObj)) {
        SC_setPdfCropPreference(cropMode);
    }
    placementSetImportPageNumber(pageNum);

    var item = doc.placedItems.add();
    item.file = fileObj;
    item.position = pos;
    placementApplyMaskToArtboard(doc, item, abRect, 0);
    return item;
}

function placementPlaceSinglePage(doc, fileObj, pageNum, pageW, pageH, nextX, baseTop, activeIdx, abCount, cropMode, abGap) {
    var singleRect = [nextX, baseTop, nextX + pageW, baseTop - pageH];
    abCount = placementUseOrAddArtboard(doc, activeIdx, singleRect, abCount);
    placementPlacePageWithMask(doc, fileObj, pageNum, singleRect, [nextX, baseTop], cropMode);
    return {
        nextX: nextX + pageW + abGap,
        abCount: abCount
    };
}

function placementPlaceSpreadHalf(doc, fileObj, pageNum, abRect, posX, baseTop, activeIdx, abCount, cropMode) {
    abCount = placementUseOrAddArtboard(doc, activeIdx, abRect, abCount);
    placementPlacePageWithMask(doc, fileObj, pageNum, abRect, [posX, baseTop], cropMode);
    return abCount;
}

function placementPlaceSpreadPage(doc, fileObj, pageNum, pageW, pageH, nextX, baseTop, activeIdx, abCount, cropMode, isR2L, abGap) {
    var halfW = pageW / 2;

    var firstRect = [nextX, baseTop, nextX + halfW, baseTop - pageH];
    var firstPosX = isR2L ? (nextX - halfW) : nextX;
    abCount = placementPlaceSpreadHalf(doc, fileObj, pageNum, firstRect, firstPosX, baseTop, activeIdx, abCount, cropMode);
    nextX += halfW + abGap;

    var secondRect = [nextX, baseTop, nextX + halfW, baseTop - pageH];
    var secondPosX = isR2L ? nextX : (nextX - halfW);
    abCount = placementPlaceSpreadHalf(doc, fileObj, pageNum, secondRect, secondPosX, baseTop, activeIdx, abCount, cropMode);
    nextX += halfW + abGap;

    return {
        nextX: nextX,
        abCount: abCount
    };
}

function placementFitAllArtboardsInView(targetDoc) {
    if (!targetDoc || !targetDoc.artboards || targetDoc.artboards.length === 0) return;

    try {
        app.activeDocument = targetDoc;
    } catch (_) { }

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
    } catch (_) { }
}


// =========================================
// PDF配置時のトリミング指定
// Art / Trim / Bleed / Crop に対応
// =========================================

// UI用トリミング定数
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

function getLastPageFromSelection(selectionItems) {
    var placed = pageCountFindFirstPlacedItem(selectionItems);
    if (!placed) return null;

    var f = placed.file;
    if (!f) {
        SC_alert('alertLinkUnknown');
        return null;
    }

    var name = decodeURIComponent(f.name);
    if (!/\.(?:pdf|ai)$/i.test(name)) return null;

    var last = getPageLengthFromFile(f);
    if (!last || isNaN(Number(last)) || Number(last) <= 0) {
        SC_alert('alertPageCountFail');
        return null;
    }

    return Number(last);
}

function updatePageCountFromPlacedOrFile(doc, fileObjOrNull, setPathTextFn, setResultTextFn) {
    var placedTemp = null;

    try {
        // ファイル指定時は一時配置して、既存の選択ベースのページ数取得ロジックを再利用
        if (fileObjOrNull) {
            var name = decodeURIComponent(fileObjOrNull.name);
            if (!/\.(?:pdf|ai)$/i.test(name)) {
                SC_alert('alertPickPdfAi');
                if (setResultTextFn) setResultTextFn(null);
                return;
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
                } catch (_) { }

                // 既存ロジックをそのまま使って総ページ数を取得
                var lastFromFile = getLastPageFromSelection([placedTemp]);
                if (setPathTextFn) setPathTextFn(fileObjOrNull);
                if (setResultTextFn) setResultTextFn(lastFromFile);
            } finally {
                if (placedTemp) {
                    try { placedTemp.remove(); } catch (_) { }
                }
            }
            return;
        }

        // ファイル未指定時は現在の選択から取得（PlacedItem が選択されている前提）
        var last = getLastPageFromSelection(doc.selection);

        // 選択中の配置画像があれば、そのファイルパス表示も更新
        try {
            var placedSel = pageCountFindFirstPlacedItem(doc.selection);
            if (placedSel && placedSel.file && setPathTextFn) setPathTextFn(placedSel.file);
        } catch (_) { }

        if (setResultTextFn) setResultTextFn(last);
    } catch (e) {
        SC_alert("alertPageCountFail", e);
        try { if (setResultTextFn) setResultTextFn(null); } catch (_) { }
    }
}

function pageCountFindFirstPlacedItem(items) {
    if (!items || items.length <= 0) return null;
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (!it) continue;
        var n = (it.constructor && it.constructor.name) ? it.constructor.name : '';
        if (n === 'PlacedItem') return it;
        if (n === 'GroupItem') {
            var hit = pageCountFindFirstPlacedItem(it.pageItems);
            if (hit) return hit;
        }
    }
    return null;
}

// pdf/aiから総ページ数を推定（既存ロジックを維持しつつ最小化）
function getPageLengthFromFile(file) {
    var tg = /<<\/Count\s(\d+)/;
    var c1 = /<<\/Type\/Page\/Parent/;
    var c2 = /\/Type\s\/Page\s/;
    var c3 = /\/StructParents\s\d+.*\/Type\/Page>>/;
    var tg2 = /<<\/Linearized\s.+\/N\s(\d+)\/T\s.+>>/;
    var tg3 = /\/Type\/Pages/;
    var tg4 = /\/Count\s(\d+)/;

    var res, wd, len = 0, num = 0;
    try {
        file.open('r');
        while (!file.eof) {
            wd = file.readln();
            if (tg.test(wd) || tg2.test(wd)) { res = Number(RegExp.$1); break; }
            if (c1.test(wd) || c2.test(wd) || c3.test(wd)) len++;
            if (tg3.test(wd)) {
                wd = file.readln();
                if (tg4.test(wd)) { num = Number(RegExp.$1); if (len < num) len = num; }
            }
        }
        if (len > 0) res = len;
    } catch (e) {
        SC_alert("alertPageCountFail", e);
    } finally {
        try { file.close(); } catch (_) { }
    }
    return res;
}

function SC_setPdfCropPreference(cropVal) {
    var keys = [
        "plugin/PDFImport/CropToBox",
        "plugin/PDFImport/CropTo",
        "plugin/PDFImport/CropBox",
        "plugin/PDFImport/CropToType"
    ];
    for (var i = 0; i < keys.length; i++) {
        try {
            app.preferences.setIntegerPreference(keys[i], cropVal);
        } catch (e) { }
    }
}

// =========================================
// PDF綴じ方向の自動検出
// /Direction /R2L があれば偶数ページを右、それ以外は左として判定
// =========================================

function SC_detectPdfBindingDirection(file) {
    // returns "R2L" or "L2R"
    var dir = "L2R";
    try {
        file.open('r');
        var lineCount = 0;
        while (!file.eof && lineCount < 200) {
            var line = file.readln();
            lineCount++;
            if (/\/Direction\s*\/R2L/.test(line)) {
                dir = "R2L";
                break;
            }
        }
    } catch (e) {
    } finally {
        try { file.close(); } catch (_) { }
    }
    return dir;
}

/**
 * Illustrator のバージョン差異を吸収するため、複数キーに試行します。
 * 期待する値（多くの環境で）: 0=Media, 1=Crop, 2=Bleed, 3=Trim, 4=Art
 */

function isPdfLikeFile(f) {
    var n = String((f && f.name) || "").toLowerCase();
    return (n.indexOf(".pdf") > -1) || (n.indexOf(".ai") > -1);
}

// トリミング設定は PDF のときのみ有効。AI では無効のままにする。
function isPdfFile(f) {
    var n = String((f && f.name) || "").toLowerCase();
    return (n.indexOf(".pdf") > -1);
}

// 入力された文字列（例："1-20", "1,3,5"）を数字の配列に変換する関数
function parsePageNumbers(inputStr) {
    var result = [];
    var parts = inputStr.split(',');

    for (var i = 0; i < parts.length; i++) {
        var part = parts[i].replace(/^\s+|\s+$/g, '');

        if (part.indexOf('-') > -1) {
            var bounds = part.split('-');
            var start = parseInt(bounds[0], 10);
            var end = parseInt(bounds[1], 10);

            if (!isNaN(start) && !isNaN(end)) {
                var min = Math.min(start, end);
                var max = Math.max(start, end);
                for (var j = min; j <= max; j++) {
                    result.push(j);
                }
            }
        } else {
            var num = parseInt(part, 10);
            if (!isNaN(num)) {
                result.push(num);
            }
        }
    }
    return result;
}

function main() {

    if (app.documents.length === 0) {
        SC_alert("alertNeedDoc");
        return;
    }

    var srcDoc = app.activeDocument;
    var docColorSpace = srcDoc.documentColorSpace;

    function getInitialOutputDocSize(srcDoc, fileObj, rangeText, cropMode) {
        var pages = parsePageNumbers(rangeText || '');
        var firstPage = (pages && pages.length > 0) ? parseInt(pages[0], 10) : 1;
        if (isNaN(firstPage) || firstPage < 1) firstPage = 1;

        if (!fileObj) {
            return {
                width: srcDoc.width,
                height: srcDoc.height
            };
        }

        try {
            // Ensure we use placementMeasurePlacedPageSize, not any old/legacy measurePlacedPageSize
            return placementMeasurePlacedPageSize(srcDoc, fileObj, firstPage, cropMode);
        } catch (_) {
            return {
                width: srcDoc.width,
                height: srcDoc.height
            };
        }
    }

    var doc = null;

    // 現在の読み込み対象ファイル。未選択時は null。
    var fileA = null;

    // ------------------------
    // UI構築
    // ------------------------
    function buildDialogUI() {
        var win = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        win.alignChildren = "fill";

        var cols = win.add("group");
        cols.orientation = "row";
        cols.alignChildren = ["fill", "top"];

        var leftCol = cols.add("group");
        leftCol.orientation = "column";
        leftCol.alignChildren = "fill";

        var rightCol = cols.add("group");
        rightCol.orientation = "column";
        rightCol.alignChildren = "fill";

        var pnlSource = leftCol.add('panel', undefined, L('panelSource'));
        pnlSource.orientation = 'column';
        pnlSource.alignChildren = ['left', 'top'];
        pnlSource.margins = [15, 20, 15, 10];
        var btnBrowse = pnlSource.add('button', undefined, L('btnLoad'));
        var etPath = pnlSource.add('statictext', undefined, L('notSelected'));
        etPath.characters = 16;

        var pnlNewDoc = rightCol.add('panel', undefined, L('panelNewDoc'));
        pnlNewDoc.orientation = 'column';
        pnlNewDoc.alignChildren = ['left', 'top'];
        pnlNewDoc.margins = [15, 20, 15, 10];
        pnlNewDoc.add('statictext', undefined, L('colorMode'));
        var rowColorMode = pnlNewDoc.add('group');
        rowColorMode.orientation = 'row';
        rowColorMode.alignChildren = ['left', 'center'];
        var rbColorCMYK = rowColorMode.add('radiobutton', undefined, L('colorModeCMYK'));
        var rbColorRGB = rowColorMode.add('radiobutton', undefined, L('colorModeRGB'));
        if (srcDoc.documentColorSpace === DocumentColorSpace.RGB) rbColorRGB.value = true;
        else rbColorCMYK.value = true;

        var pnlAB = leftCol.add('panel', undefined, L('panelLoad'));
        pnlAB.orientation = 'column';
        pnlAB.alignChildren = ['left', 'top'];
        pnlAB.margins = [15, 20, 15, 10];
        var rowRange = pnlAB.add('group');
        rowRange.orientation = 'row';
        rowRange.alignChildren = ['left', 'center'];
        rowRange.add('statictext', undefined, L('range'));
        var etRange = rowRange.add('edittext', undefined, '');
        etRange.characters = 10;
        etRange.enabled = true;

        var panelCrop = rightCol.add("panel", undefined, L("panelItem"));
        panelCrop.alignChildren = "left";
        panelCrop.margins = [15, 20, 15, 10];
        var ddCrop = panelCrop.add("dropdownlist", undefined, [L("cropArt"), L("cropTrim"), L("cropCrop"), L("cropBleed")]);
        ddCrop.minimumSize.width = 160;
        ddCrop.selection = 2;
        var groupBind = panelCrop.add("group");
        groupBind.orientation = "row";
        groupBind.alignChildren = ["left", "center"];
        groupBind.add("statictext", undefined, L("evenPage"));
        var rbR2L = groupBind.add("radiobutton", undefined, L("bindR2L"));
        var rbL2R = groupBind.add("radiobutton", undefined, L("bindL2R"));
        rbR2L.value = true;

        var groupButtons = win.add("group");
        groupButtons.orientation = "row";
        groupButtons.alignChildren = ["center", "center"];
        var btnCancel = groupButtons.add("button", undefined, L("cancel"), { name: "cancel" });
        var btnOk = groupButtons.add("button", undefined, L("ok"), { name: "ok" });

        return {
            win: win,
            btnBrowse: btnBrowse,
            etPath: etPath,
            rbColorCMYK: rbColorCMYK,
            rbColorRGB: rbColorRGB,
            etRange: etRange,
            ddCrop: ddCrop,
            rbR2L: rbR2L,
            rbL2R: rbL2R,
            btnCancel: btnCancel,
            btnOk: btnOk
        };
    }

    // ダイアログUIを構築し、必要な参照を取り出す
    var ui = buildDialogUI();
    var win = ui.win;
    var btnBrowse = ui.btnBrowse;
    var etPath = ui.etPath;
    var rbColorCMYK = ui.rbColorCMYK;
    var rbColorRGB = ui.rbColorRGB;
    var etRange = ui.etRange;
    var ddCrop = ui.ddCrop;
    var rbR2L = ui.rbR2L;
    var rbL2R = ui.rbL2R;
    var btnCancel = ui.btnCancel;
    var btnOk = ui.btnOk;


    // ------------------------
    // UI表示補助
    // ------------------------
    function getFileDisplayName(f) {
        if (!f) return L('notSelected');
        try {
            return decodeURIComponent(f.name);
        } catch (_) {
            return String(f.name || L('notSelected'));
        }
    }

    function getFileDisplayPath(f) {
        if (!f) return '';
        try {
            return decodeURIComponent(f.fsName);
        } catch (_) {
            return String(f.fsName || '');
        }
    }

    function setPathText(f) {
        // 現在の読み込み対象ファイル参照も更新
        fileA = f || null;

        // トリミング設定は PDF のときのみ有効化
        ddCrop.enabled = isPdfFile(fileA);

        // PDF の /Direction を見て偶数ページ位置を自動設定
        autoDetectBinding(f);

        etPath.text = getFileDisplayName(f);
        etPath.helpTip = getFileDisplayPath(f);
    }

    function setResultText(last) {
        etRange.text = last ? ('1-' + last) : '';
    }


    // ------------------------
    // UIイベント配線
    // ------------------------
    function bindDialogEvents() {
        btnBrowse.onClick = function () {
            var f = File.openDialog(L('dlgPickFile'), L('filterPick'));
            if (!f) return;
            updatePageCountFromPlacedOrFile(srcDoc, f, setPathText, setResultText);
        };

        btnCancel.onClick = function () {
            closeDialog(2);
        };

        btnOk.onClick = function () {
            closeDialog(1);
        };
    }

    // ------------------------
    // 初期状態反映
    // ------------------------
    function initializeDialogState() {
        ddCrop.enabled = isPdfFile(fileA);
        updatePageCountFromPlacedOrFile(srcDoc, null, setPathText, setResultText);
        ddCrop.enabled = isPdfFile(fileA);
    }

    // ------------------------
    // 出力ドキュメント設定
    // ------------------------
    function getSelectedDocColorSpace() {
        return rbColorRGB.value ? DocumentColorSpace.RGB : DocumentColorSpace.CMYK;
    }

    // 現在は UI から変更せず、既定値を返す
    function getArtboardGap() {
        return DEFAULT_ARTBOARD_GAP;
    }

    // 現在は UI から変更せず、既定値を返す
    function getRasterEffectsResolution() {
        return DEFAULT_RASTER_EFFECTS_RESOLUTION;
    }

    function applyDocumentRasterEffectsResolution(targetDoc) {
        if (!targetDoc) return;
        try {
            var res = getRasterEffectsResolution();
            var settings = targetDoc.rasterEffectSettings;
            settings.resolution = res;
            targetDoc.rasterEffectSettings = settings;
        } catch (_) { }
    }


    // ------------------------
    // 配置オプション取得
    // ------------------------
    function isR2L() {
        return !!rbR2L.value;
    }

    // ファイル選択時に綴じ方向を自動検出して反映
    function autoDetectBinding(f) {
        if (!f) return;
        try {
            var dir = SC_detectPdfBindingDirection(f);
            if (dir === "R2L") {
                rbR2L.value = true;
            } else {
                rbL2R.value = true;
            }
        } catch (_) { }
    }

    function getCropModeFromUI() {
        var idx = (ddCrop.selection) ? ddCrop.selection.index : 2;
        // 0:アート / 1:トリミング / 2:仕上がり / 3:裁ち落とし
        if (idx === 0) return CROP_ART;
        if (idx === 1) return CROP_TRIM;
        if (idx === 3) return CROP_BLEED;
        // 仕上がりは CropBox を想定
        return CROP_CROP;
    }

    // ------------------------
    // 実行処理
    // ------------------------
    function placeOnIndividualArtboards(targetPages, cropMode, isR2L) {
        if (!doc) {
            var initialSize = getInitialOutputDocSize(srcDoc, fileA, etRange.text, cropMode);
            docColorSpace = getSelectedDocColorSpace();
            doc = app.documents.add(
                docColorSpace,
                initialSize.width,
                initialSize.height
            );
            applyDocumentRasterEffectsResolution(doc);
            app.activeDocument = doc;
        }
        var activeIdx = doc.artboards.getActiveArtboardIndex();
        var baseRect = doc.artboards[activeIdx].artboardRect;
        var nextX = baseRect[0];
        var baseTop = baseRect[1];

        var abGap = getArtboardGap();
        var abCount = 0;
        var completed = false;

        try {
            for (var i = 0; i < targetPages.length; i++) {
                var pageNum = parseInt(targetPages[i], 10);
                if (isNaN(pageNum) || pageNum < 1) pageNum = 1;

                var pageSize = placementMeasurePlacedPageSize(doc, fileA, pageNum, cropMode);
                var pageW = pageSize.width;
                var pageH = pageSize.height;
                var isSpread = (pageW > pageH * 1.2);
                var result;

                if (isSpread) {
                    result = placementPlaceSpreadPage(doc, fileA, pageNum, pageW, pageH, nextX, baseTop, activeIdx, abCount, cropMode, isR2L, abGap);
                } else {
                    result = placementPlaceSinglePage(doc, fileA, pageNum, pageW, pageH, nextX, baseTop, activeIdx, abCount, cropMode, abGap);
                }

                nextX = result.nextX;
                abCount = result.abCount;
            }
            completed = true;
        } catch (e) {
            SC_alert("alertPlaceError", e);
        } finally {
            try { placementResetImportPageNumber(); } catch (_) { }
        }

        if (completed) {
            placementFitAllArtboardsInView(doc);
        }
    }

    // ------------------------
    // ダイアログ終了処理
    // ------------------------
    function closeDialog(resultCode) {
        win.close(resultCode);
    }

    bindDialogEvents();
    initializeDialogState();

    if (win.show() === 1) {
        if (!fileA) {
            SC_alert('alertNeedFile');
        } else {
            var finalPages = parsePageNumbers(etRange.text);
            if (!finalPages || finalPages.length === 0) finalPages = [1];

            var cropMode = getCropModeFromUI();
            placeOnIndividualArtboards(finalPages, cropMode, isR2L());
        }
    }

}

main();