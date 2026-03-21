#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * SlideCollage-mihiraki.jsx
 *
 * 概要:
 * PDF/AIファイルを指定したページ範囲で配置し、各ページを個別のアートボードとして展開します。
 * 横長ページは見開きとみなして左右に分割し、綴じ方向（右綴じ/左綴じ）に応じた順序で配置できます。
 * PDFではトリミング設定を選択でき、選択中の配置画像または指定ファイルからページ数を推定して
 * ページ範囲欄へ反映します。ダイアログ位置はセッション中のみ記憶します。
 *
 * 更新日: 2026-03-17
 */

// =========================================
// Version & Localization
// =========================================

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "スライドコラージュ（見開き対応）",
        en: "Slide Collage (Spread Support)"
    },

    // Panels
    panelLoad: { ja: "アートボード", en: "Artboards" },
    panelSource: { ja: "読み込みファイル", en: "Source File" },
    panelItem: { ja: "アイテム", en: "Items" },

    // Load panel
    btnLoad: { ja: "ファイル指定", en: "Select File" },
    range: { ja: "指定", en: "Range" },
    dlgPickFile: { ja: "PDF/AIを選択してください", en: "Select a PDF/AI" },
    filterPick: { ja: "PDF/AI:*.pdf;*.ai", en: "PDF/AI:*.pdf;*.ai" },
    alertLinkUnknown: { ja: "画像のリンク先が不明でした。", en: "Image link not found." },
    alertPageCountFail: { ja: "リンクされたPDF/AIファイルのページ数を取得できませんでした。", en: "Could not determine the page count of the linked PDF/AI file." },
    alertPickPdfAi: { ja: "PDFまたはAIファイルを選択してください。", en: "Please select a PDF or AI file." },
    notSelected: { ja: "未指定", en: "Not selected" },

    // Binding direction
    bindR2L: { ja: "右綴じ", en: "Right-to-Left" },
    bindL2R: { ja: "左綴じ", en: "Left-to-Right" },

    // Item panel
    cropArt: { ja: "アート", en: "Art" },
    cropTrim: { ja: "トリミング", en: "Trim" },
    cropCrop: { ja: "仕上がり", en: "Crop" },
    cropBleed: { ja: "裁ち落とし", en: "Bleed" },

    // Buttons
    cancel: { ja: "キャンセル", en: "Cancel" },
    ok: { ja: "OK", en: "OK" },

    // File / Alerts
    alertNeedDoc: { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document before running." },
    alertPlaceError: { ja: "配置中にエラーが発生しました。", en: "An error occurred while placing the pages." },
    alertNeedFile: { ja: "先に［ファイル指定］で読み込みファイルを選択してください。", en: "Please select a source file first." },

};

function L(key) {
    var o = LABELS[key];
    if (!o) return key;
    return o[lang] || o.en || o.ja || key;
}

// Safe alert helper (used by __TMKPageCount_ module)
if (typeof safeAlertKey === "undefined") {
    var safeAlertKey = function (key) {
        try { alert(L(key)); } catch (_) { }
    };
}

// =========================================
// PDF Import crop (trim / bleed / art)
// =========================================

// UI mode constants
var __SC_CROP_ART = 4;
var __SC_CROP_TRIM = 3;
var __SC_CROP_BLEED = 2;
var __SC_CROP_CROP = 1;


// ============================================================
// TMK Page Count Module (collision-safe)
// - Estimate total pages (last page number) of linked PDF/AI
// - No relink / no placedItems.add
// ============================================================

function __TMKPageCount_getLastPageFromSelection(selectionItems) {
    var placed = __TMKPageCount_findFirstPlacedItem(selectionItems);
    if (!placed) return null;

    var f = placed.file;
    if (!f) {
        safeAlertKey('alertLinkUnknown');
        return null;
    }

    var name = decodeURIComponent(f.name);
    if (!/\.(?:pdf|ai)$/i.test(name)) return null;

    var last = __TMKPageCount_getPageLengthFromFile(f);
    if (!last || isNaN(Number(last)) || Number(last) <= 0) {
        safeAlertKey('alertPageCountFail');
        return null;
    }

    return Number(last);
}

function __TMKPageCount_updateResultFromPlacedOrFile(doc, fileObjOrNull, setPathTextFn, setResultTextFn) {
    var placedTemp = null;

    try {
        // If a file is provided, place it temporarily to reuse existing selection-based logic
        if (fileObjOrNull) {
            var name = decodeURIComponent(fileObjOrNull.name);
            if (!/\.(?:pdf|ai)$/i.test(name)) {
                safeAlertKey('alertPickPdfAi');
                if (setResultTextFn) setResultTextFn(null);
                return;
            }

            try {
                placedTemp = doc.placedItems.add();
                placedTemp.file = fileObjOrNull;

                // Place near the visible top-left so it can be seen (briefly)
                try {
                    var vb = doc.activeView && doc.activeView.bounds ? doc.activeView.bounds : null; // [left, top, right, bottom]
                    if (vb && vb.length === 4) {
                        placedTemp.position = [vb[0], vb[1]];
                    }
                } catch (_) { }

                // Use the existing logic without modifying it
                var lastFromFile = __TMKPageCount_getLastPageFromSelection([placedTemp]);
                if (setPathTextFn) setPathTextFn(fileObjOrNull);
                if (setResultTextFn) setResultTextFn(lastFromFile);
            } finally {
                if (placedTemp) {
                    try { placedTemp.remove(); } catch (_) { }
                }
            }
            return;
        }

        // No file provided: use current selection (PlacedItem must be selected)
        var last = __TMKPageCount_getLastPageFromSelection(doc.selection);

        // Attempt to show the currently selected placed file path
        try {
            var placedSel = __TMKPageCount_findFirstPlacedItem(doc.selection);
            if (placedSel && placedSel.file && setPathTextFn) setPathTextFn(placedSel.file);
        } catch (_) { }

        if (setResultTextFn) setResultTextFn(last);
    } catch (e) {
        alert(e);
        try { if (setResultTextFn) setResultTextFn(null); } catch (_) { }
    }
}

function __TMKPageCount_findFirstPlacedItem(items) {
    if (!items || items.length <= 0) return null;
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (!it) continue;
        var n = (it.constructor && it.constructor.name) ? it.constructor.name : '';
        if (n === 'PlacedItem') return it;
        if (n === 'GroupItem') {
            var hit = __TMKPageCount_findFirstPlacedItem(it.pageItems);
            if (hit) return hit;
        }
    }
    return null;
}

// pdf/aiから総ページ数を推定（既存ロジックを維持しつつ最小化）
function __TMKPageCount_getPageLengthFromFile(file) {
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
        alert(e);
    } finally {
        try { file.close(); } catch (_) { }
    }
    return res;
}

/**
 * Set PDF import crop box preference.
 * Illustrator のバージョン差異を吸収するため、複数キーに試行します。
 * 期待する値（多くの環境で）: 0=Media, 1=Crop, 2=Bleed, 3=Trim, 4=Art
 */
function __SC_setPdfCropPreference(cropVal) {
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
// /Direction /R2L があれば右綴じ、なければ左綴じと判定
// =========================================
function __SC_detectPdfBindingDirection(file) {
    // returns "R2L" or "L2R"
    try {
        file.open('r');
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
        try { file.close(); } catch (_) { }
    }
    return "L2R";
}

// Placing .ai in Illustrator uses the PDF import pipeline as well (AI is PDF-compatible),
// so we treat .ai as "PDF-like" for page/artboard selection.
function __SC_isPdfLikeFile(f) {
    try {
        if (!f) return false;
        var n = (f.name || "").toLowerCase();
        return (n.indexOf(".pdf") > -1) || (n.indexOf(".ai") > -1);
    } catch (e) {
        return false;
    }
}

// UI-only: crop box options are meaningful for PDF; keep disabled for AI.
function __SC_isPdfFile(f) {
    try {
        if (!f) return false;
        var n = (f.name || "").toLowerCase();
        return (n.indexOf(".pdf") > -1);
    } catch (e) {
        return false;
    }
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

// =========================================
// Dialog position (session only)
// =========================================

function __SC_getSessionDialogBoundsKey() {
    return "__SC_slideCollage_dialog_bounds";
}

function __SC_loadDialogBounds() {
    try {
        var k = __SC_getSessionDialogBoundsKey();
        return $.global[k];
    } catch (e) {
        return null;
    }
}

function __SC_saveDialogBounds(b) {
    try {
        if (!b) return;
        // store plain object (avoid ScriptUI objects)
        var l = b.left, t = b.top, r = b.right, bt = b.bottom;
        if (!(typeof l === "number" && isFinite(l))) return;
        if (!(typeof t === "number" && isFinite(t))) return;
        if (!(typeof r === "number" && isFinite(r))) return;
        if (!(typeof bt === "number" && isFinite(bt))) return;
        var o = { left: l, top: t, right: r, bottom: bt };
        $.global[__SC_getSessionDialogBoundsKey()] = o;
    } catch (e) { }
}


function __SC_applyDialogBounds(win) {
    try {
        var b = __SC_loadDialogBounds();
        if (!b) return;

        function isFiniteNum(n) {
            return (typeof n === "number") && isFinite(n);
        }

        var l = b.left, t = b.top, r = b.right, bt = b.bottom;
        if (!isFiniteNum(l) || !isFiniteNum(t) || !isFiniteNum(r) || !isFiniteNum(bt)) return;

        // Size sanity
        var w = r - l;
        var h = bt - t;
        if (!(w >= 200 && h >= 120)) return; // too small/invalid -> skip

        // Screen sanity: must intersect at least one screen bounds
        var intersects = false;
        try {
            if ($.screens && $.screens.length) {
                for (var i = 0; i < $.screens.length; i++) {
                    var sb = $.screens[i].bounds; // [L,T,R,B]
                    var sl = sb[0], st = sb[1], sr = sb[2], sbt = sb[3];
                    var il = Math.max(l, sl);
                    var it = Math.max(t, st);
                    var ir = Math.min(r, sr);
                    var ib = Math.min(bt, sbt);
                    if ((ir - il) > 40 && (ib - it) > 40) { // require some visible area
                        intersects = true;
                        break;
                    }
                }
            }
        } catch (eScr) {
            // If screens API is unavailable, fall back to applying.
            intersects = true;
        }

        if (!intersects) return;

        // Apply bounds
        win.bounds = [l, t, r, bt];
    } catch (e) { }
}


function main() {


    if (app.documents.length === 0) {
        alert(L("alertNeedDoc"));
        return;
    }

    var doc = app.activeDocument;
    // Current source file selected by user (set via UI). If null, the user has not selected any file yet.
    var fileA = null;




    // 2. ダイアログボックスの作成
    var win = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    win.alignChildren = "fill";

    // セッション中のダイアログ位置を復元 / Restore dialog position in this session
    __SC_applyDialogBounds(win);

    // If no valid stored bounds, center as a safe fallback
    try {
        if (!__SC_loadDialogBounds()) {
            win.center();
        }
    } catch (eCenter) { }

    // 移動時に位置を記憶 / Remember position while moving
    win.onMove = function () { __SC_saveDialogBounds(win.bounds); };
    win.onMoving = function () { __SC_saveDialogBounds(win.bounds); };

    // === 単カラム構成 ===
    var leftCol = win.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = "fill";

    // Panel: 読み込みファイル / Source file
    var pnlSource = leftCol.add('panel', undefined, L('panelSource'));
    pnlSource.orientation = 'column';
    pnlSource.alignChildren = ['left', 'top'];
    pnlSource.margins = [15, 20, 15, 10];

    // ファイル指定ボタン（fileA とページ範囲表示を更新）
    var btnBrowse = pnlSource.add('button', undefined, L('btnLoad'));

    // ファイル名表示
    var etPath = pnlSource.add('statictext', undefined, L('notSelected'));
    etPath.characters = 20;

    // Panel: アートボード
    var pnlAB = leftCol.add('panel', undefined, L('panelLoad'));
    pnlAB.orientation = 'column';
    pnlAB.alignChildren = ['left', 'top'];
    pnlAB.margins = [15, 20, 15, 10];

    // 範囲
    var rowRange = pnlAB.add('group');
    rowRange.orientation = 'row';
    rowRange.alignChildren = ['left', 'center'];
    rowRange.add('statictext', undefined, L('range'));
    var etRange = rowRange.add('edittext', undefined, '');
    etRange.characters = 10;
    etRange.enabled = true;



    function setPathText(f) {
        // Also update current source file reference (fileA)
        fileA = f || null;

        // Update crop dropdown availability (crop is meaningful only for PDF)
        try { ddCrop.enabled = __SC_isPdfFile(fileA); } catch (_) { }

        // 綴じ方向を自動検出
        try { __SC_autoDetectBinding(f); } catch (_) { }

        try {
            if (f) {
                etPath.text = decodeURIComponent(f.name);
                etPath.helpTip = decodeURIComponent(f.fsName);
            } else {
                etPath.text = L('notSelected');
                etPath.helpTip = '';
            }
        } catch (_) {
            try {
                etPath.text = f ? String(f.name) : L('notSelected');
                etPath.helpTip = f ? String(f.fsName) : '';
            } catch (__) {
                etPath.text = L('notSelected');
                etPath.helpTip = '';
            }
        }
    }

    function setResultText(last) {
        if (!last) {
            try { etRange.text = ''; } catch (_) { }
        } else {
            try { etRange.text = '1-' + last; } catch (_) { }
        }
    }

    // Initial: try selection
    __TMKPageCount_updateResultFromPlacedOrFile(doc, null, setPathText, setResultText);
    try { ddCrop.enabled = __SC_isPdfFile(fileA); } catch (_) { }

    btnBrowse.onClick = function () {
        var f = File.openDialog(L('dlgPickFile'), L('filterPick'));
        if (!f) return;
        __TMKPageCount_updateResultFromPlacedOrFile(doc, f, setPathText, setResultText);
    };


    // --- アイテムパネル（PDF のトリミング設定と綴じ方向） ---
    var panelCrop = leftCol.add("panel", undefined, L("panelItem"));
    panelCrop.alignChildren = "left";
    panelCrop.margins = [15, 20, 15, 10];

    // panelCrop.add("statictext", undefined, "トリミング");
    var ddCrop = panelCrop.add("dropdownlist", undefined, [L("cropArt"), L("cropTrim"), L("cropCrop"), L("cropBleed")]);
    ddCrop.minimumSize.width = 160;

    // デフォルト：仕上がり
    ddCrop.selection = 2;

    ddCrop.enabled = __SC_isPdfFile(fileA); // fileA may be null until a file is selected

    // 綴じ方向（右綴じ / 左綴じ）
    var groupBind = panelCrop.add("group");
    groupBind.orientation = "row";
    groupBind.alignChildren = ["left", "center"];

    var rbR2L = groupBind.add("radiobutton", undefined, L("bindR2L"));
    var rbL2R = groupBind.add("radiobutton", undefined, L("bindL2R"));
    rbR2L.value = true; // デフォルト：右綴じ

    function __SC_isR2L() {
        return !!rbR2L.value;
    }

    // ファイル選択時に自動検出して反映
    function __SC_autoDetectBinding(f) {
        if (!f) return;
        try {
            var dir = __SC_detectPdfBindingDirection(f);
            if (dir === "R2L") {
                rbR2L.value = true;
            } else {
                rbL2R.value = true;
            }
        } catch (_) { }
    }

    function __SC_getCropModeFromUI() {
        var idx = (ddCrop.selection) ? ddCrop.selection.index : 2;
        // 0:アート / 1:トリミング / 2:仕上がり / 3:裁ち落とし
        if (idx === 0) return __SC_CROP_ART;
        if (idx === 1) return __SC_CROP_TRIM;
        if (idx === 3) return __SC_CROP_BLEED;
        // 仕上がりは CropBox を想定（環境差があるため失敗時は無視される）
        return __SC_CROP_CROP;
    }

    // -----------------------------------------
    // Import page/artboard selection
    // - PDF: uses PDFImport/PageNumber
    // - AI: uses IllustratorImport/ArtboardNumber (+ PlaceArtboards)
    // -----------------------------------------
    function __SC_setImportPageNumber(pageNum) {
        var n = parseInt(pageNum, 10);
        if (isNaN(n) || n < 1) n = 1;
        try {
            // Use PDFImport for both PDF and AI (AI is handled by the PDF import pipeline when placing)
            app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", n);
        } catch (_) { }
    }

    function __SC_resetImportPageNumber() {
        try {
            app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", 1);
        } catch (_) { }
    }


    // ボタン類（キャンセル・OK）
    var groupButtons = win.add("group");
    groupButtons.orientation = "row";
    groupButtons.alignChildren = ["center", "center"];

    var btnCancel = groupButtons.add("button", undefined, L("cancel"), { name: "cancel" });
    var btnOk = groupButtons.add("button", undefined, L("ok"), { name: "ok" });

    // -----------------------------------------
    // 指定アートボードのマージン内側でクリッピングマスクを適用
    // -----------------------------------------
    function __SC_applyMaskToArtboard(doc, item, abRect, marginPt) {
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

    // -----------------------------------------
    // Helper functions for placing pages and artboards
    // -----------------------------------------

    function __SC_measurePlacedPageSize(doc, fileObj, pageNum, cropMode) {
        if (__SC_isPdfLikeFile(fileObj)) {
            __SC_setPdfCropPreference(cropMode);
        }
        __SC_setImportPageNumber(pageNum);

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

    function __SC_useOrAddArtboard(doc, activeIdx, abRect, abCount) {
        if (abCount === 0) {
            doc.artboards[activeIdx].artboardRect = abRect;
        } else {
            doc.artboards.add(abRect);
        }
        return abCount + 1;
    }

    function __SC_placePageWithMask(doc, fileObj, pageNum, abRect, pos, cropMode) {
        if (__SC_isPdfLikeFile(fileObj)) {
            __SC_setPdfCropPreference(cropMode);
        }
        __SC_setImportPageNumber(pageNum);

        var item = doc.placedItems.add();
        item.file = fileObj;
        item.position = pos;
        __SC_applyMaskToArtboard(doc, item, abRect, 0);
        return item;
    }

    function __SC_placeSinglePage(doc, fileObj, pageNum, pageW, pageH, nextX, baseTop, activeIdx, abCount, cropMode) {
        var singleRect = [nextX, baseTop, nextX + pageW, baseTop - pageH];
        abCount = __SC_useOrAddArtboard(doc, activeIdx, singleRect, abCount);

        __SC_placePageWithMask(doc, fileObj, pageNum, singleRect, [nextX, baseTop], cropMode);

        return {
            nextX: nextX + pageW + 20,
            abCount: abCount
        };
    }

    function __SC_placeSpreadHalf(doc, fileObj, pageNum, abRect, posX, baseTop, activeIdx, abCount, cropMode) {
        abCount = __SC_useOrAddArtboard(doc, activeIdx, abRect, abCount);
        __SC_placePageWithMask(doc, fileObj, pageNum, abRect, [posX, baseTop], cropMode);
        return abCount;
    }

    function __SC_placeSpreadPage(doc, fileObj, pageNum, pageW, pageH, nextX, baseTop, activeIdx, abCount, cropMode, isR2L, abGap) {
        var halfW = pageW / 2;

        var firstRect = [nextX, baseTop, nextX + halfW, baseTop - pageH];
        var firstPosX = isR2L ? (nextX - halfW) : nextX;
        abCount = __SC_placeSpreadHalf(doc, fileObj, pageNum, firstRect, firstPosX, baseTop, activeIdx, abCount, cropMode);
        nextX += halfW + abGap;

        var secondRect = [nextX, baseTop, nextX + halfW, baseTop - pageH];
        var secondPosX = isR2L ? nextX : (nextX - halfW);
        abCount = __SC_placeSpreadHalf(doc, fileObj, pageNum, secondRect, secondPosX, baseTop, activeIdx, abCount, cropMode);
        nextX += halfW + abGap;

        return {
            nextX: nextX,
            abCount: abCount
        };
    }

    // -----------------------------------------
    // 各ページを個別のアートボードに配置
    // 見開き（横長）ページは左右に分割して2つのアートボードに配置
    // -----------------------------------------
    function placeOnIndividualArtboards(targetPages, cropMode, isR2L) {
        var activeIdx = doc.artboards.getActiveArtboardIndex();
        var baseRect = doc.artboards[activeIdx].artboardRect;
        var nextX = baseRect[0];
        var baseTop = baseRect[1];

        var abGap = 20; // アートボード間の間隔（pt）
        var abCount = 0; // 作成/使用したアートボード数

        try {
            for (var i = 0; i < targetPages.length; i++) {
                var pageNum = parseInt(targetPages[i], 10);
                if (isNaN(pageNum) || pageNum < 1) pageNum = 1;

                var pageSize = __SC_measurePlacedPageSize(doc, fileA, pageNum, cropMode);
                var pageW = pageSize.width;
                var pageH = pageSize.height;

                // 見開き判定：横長（幅 > 高さ × 1.2）なら見開き
                var isSpread = (pageW > pageH * 1.2);
                var result;

                if (isSpread) {
                    result = __SC_placeSpreadPage(doc, fileA, pageNum, pageW, pageH, nextX, baseTop, activeIdx, abCount, cropMode, isR2L, abGap);
                } else {
                    result = __SC_placeSinglePage(doc, fileA, pageNum, pageW, pageH, nextX, baseTop, activeIdx, abCount, cropMode);
                }

                nextX = result.nextX;
                abCount = result.abCount;
            }
        } catch (e) {
            alert(L("alertPlaceError"));
        } finally {
            try { __SC_resetImportPageNumber(); } catch (_) { }
        }
    }

    function __SC_closeDialog(resultCode) {
        __SC_saveDialogBounds(win.bounds);
        win.close(resultCode);
    }

    btnCancel.onClick = function () {
        __SC_closeDialog(2);
    };

    btnOk.onClick = function () {
        __SC_closeDialog(1);
    };

    if (win.show() === 1) {
        if (!fileA) {
            safeAlertKey('alertNeedFile');
        } else {
            var finalPages = parsePageNumbers(etRange.text);
            if (!finalPages || finalPages.length === 0) finalPages = [1];

            var cropMode = __SC_getCropModeFromUI();
            placeOnIndividualArtboards(finalPages, cropMode, __SC_isR2L());
        }
    }

    // ダイアログ終了後に選択解除
    try { doc.selection = null; } catch (eSel) { }
}

main();