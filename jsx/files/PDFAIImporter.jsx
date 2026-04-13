#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * PDFAIImporter.jsx
 *
 * 概要:
 * PDF/AIファイルを指定したページ範囲で読み込み、現在のドキュメント上にページを配置します。
 * 配置方法は、各ページを個別のアートボードとして並べる方法と、アートボードを追加せずに
 * オブジェクトとして配置する方法から選択できます。
 *
 * 対象ページは次から選択できます。
 * ・全ページ（初期値）
 * ・先頭ページのみ
 * ・指定ページ（例: 1-10, 1,3,5）
 *
 * 読み込みファイルが確定するまでは、対象ページパネルと配置方法パネルは無効になります。
 * 選択中の配置画像、または［ファイル指定］で選んだPDF/AIファイルから総ページ数を推定し、
 * 対象ページパネルには「実際に配置されるページ数／総ページ数」を表示します。
 * 指定ページ入力欄には、検出した総ページ数をもとに既定値を保持します。
 *
 * ［ファイル指定］のダイアログでは、PDF/AIファイルのみを選択対象にします。
 *
 * 列数を指定すると、指定数ごとに改行してグリッド状に配置します。
 * 列数が自動のときは、カンバス右端に達したタイミングで折り返します。
 * 列数入力欄の右側には、列数指定時のみ推定行×列を表示します。
 *
 * ［アートボードごと］では、各ページサイズに合わせてアートボードを作成または更新し、
 * 倍率は 100% 固定で配置します。
 *
 * ［アートボードを無視］では、現在のアートボード左上を基準に、指定した倍率で
 * オブジェクトとして配置します。配置後は、新しく配置したオブジェクトを選択状態にし、
 * それら全体が見えるように表示を自動調整します。
 *
 * 間隔は配置方法に応じて意味が変わります。
 * ・アートボードごと：アートボードの間隔
 * ・アートボードを無視：配置オブジェクト同士の間隔
 *
 * ケイは次から選択できます。
 * ・なし
 * ・ケイ線のみを追加
 *
 * ケイ線のみを追加する場合は、角丸オプションを有効にして角丸半径を指定できます。
 *
 * 列数・間隔・倍率の数値入力欄では、↑↓キーで ±1、Shift+↑↓キーで ±10 の増減ができます。
 * 列数が「自動」のときは、↑キーで 1 に切り替えます。
 *
 * 配置完了後は、配置結果がすべて見えるように表示を自動調整します。
 * （アートボードごと：全アートボード表示 / 無視：配置オブジェクト全体表示）
 *
 * 更新日: 2026-04-13
 */

// =========================================
// バージョンとローカライズ
// =========================================

var SCRIPT_VERSION = "v1.1.0";

// アートボードごと配置時のデフォルト間隔（pt）
var DEFAULT_ARTBOARD_GAP = 100;

// カンバスの右端座標（220インチ / 2 × 72 pt/inch）
var CANVAS_RIGHT_EDGE = (220 / 2) * 72; // 7920 pt


function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 */

var LABELS = {
    dialogTitle: {
        ja: "PDF/AI配置",
        en: "PDF/AI Placement"
    },

    // パネル
    panelLoad: { ja: "対象ページ", en: "Pages" },
    panelSource: { ja: "読み込みファイル", en: "Source File" },
    panelItem: { ja: "配置方法", en: "Placement" },

    // 読み込みパネル
    btnLoad: { ja: "ファイル指定", en: "Select File" },
    rangeAll: { ja: "全ページ", en: "All Pages" },
    rangeFirst: { ja: "先頭ページのみ", en: "First Page Only" },
    rangeCustom: { ja: "指定ページ：", en: "Custom Pages:" },
    dlgPickFile: { ja: "PDF/AIを選択してください", en: "Select a PDF/AI" },
    alertLinkUnknown: { ja: "画像のリンク先が不明でした。", en: "Image link not found." },
    alertPageCountFail: { ja: "リンクされたPDF/AIファイルのページ数を取得できませんでした。", en: "Could not determine the page count of the linked PDF/AI file." },
    alertPickPdfAi: { ja: "PDFまたはAIファイルを選択してください。", en: "Please select a PDF or AI file." },
    notSelected: { ja: "未指定", en: "Not selected" },

    // モードパネル
    panelMode: { ja: "モード", en: "Placement Mode" },
    placePerArtboard: { ja: "アートボードごと", en: "Per Artboard" },
    placeIgnoreArtboard: { ja: "アートボードを無視", en: "Place as Objects" },
    // 配置方法パネル
    scaleLabel: { ja: "倍率:", en: "Scale" },
    scaleUnit: { ja: "%", en: "%" },
    gapLabel: { ja: "間隔:", en: "Gap" },
    gapHelpPerArtboard: { ja: "アートボードの間隔", en: "Artboard gap" },
    gapHelpIgnoreArtboard: { ja: "配置するオブジェクトの間隔", en: "Placed object gap" },
    artboardGapUnit: { ja: "pt", en: "pt" },
    colsLabel: { ja: "列数:", en: "Columns:" },
    colsAuto: { ja: "自動", en: "Auto" },
    cropArt: { ja: "アート", en: "Art" },
    cropTrim: { ja: "トリミング", en: "Trim" },
    cropCrop: { ja: "仕上がり", en: "Crop" },
    cropBleed: { ja: "裁ち落とし", en: "Bleed" },

    // ケイパネル
    panelKei: { ja: "ケイ", en: "Stroke" },
    keiNone: { ja: "なし", en: "None" },
    keiClipGroup: { ja: "クリップグループ化して追加", en: "Add Stroke Only" },

    // ケイオプション
    keiRoundCorner: { ja: "角丸", en: "Round corners" },

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

// 単位コードとラベルのマップ
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

function getCurrentRulerUnitCode() {
    try {
        return app.preferences.getIntegerPreference("rulerType");
    } catch (_) {
        return 2;
    }
}

function getCurrentUnitLabel() {
    var unitCode = getCurrentRulerUnitCode();
    return unitLabelMap[unitCode] || "pt";
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
// - アートボードの作成 / 更新
// - 単ページの配置
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

function placementPlacePage(doc, fileObj, pageNum, pos, cropMode, scalePct) {
    if (isPdfLikeFile(fileObj)) {
        SC_setPdfCropPreference(cropMode);
    }
    placementSetImportPageNumber(pageNum);

    var item = doc.placedItems.add();
    item.file = fileObj;
    item.position = pos;

    if (typeof scalePct === "number" && scalePct !== 100) {
        item.resize(scalePct, scalePct);
    }
    return item;
}

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
        } catch (_) {
            try {
                b = it.geometricBounds;
            } catch (_) { }
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

function placementFitItemsInView(targetDoc, items) {
    if (!targetDoc || !items || items.length === 0) return;

    try {
        app.activeDocument = targetDoc;
    } catch (_) { }

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
    } catch (_) { }
}


// =========================================
// ケイ処理ヘルパー
// =========================================

// 配置アイテムの外接矩形でクリッピングマスクグループを作成
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

// 角丸 LiveEffect を適用
function keiApplyRoundCornersLiveEffect(targetItem, radius) {
    if (!targetItem) return;
    var r = Number(radius);
    if (isNaN(r) || r <= 0) return;
    var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + r + ' "/></LiveEffect>';
    try {
        targetItem.applyEffect(xml);
    } catch (_) { }
}

// 配置アイテムにケイ処理を適用
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

function updatePageCountFromPlacedOrFile(doc, fileObjOrNull, setPathTextFn) {
    var placedTemp = null;

    try {
        // ファイル指定時は一時配置して、既存の選択ベースのページ数取得ロジックを再利用
        if (fileObjOrNull) {
            var name = decodeURIComponent(fileObjOrNull.name);
            if (!/\.(?:pdf|ai)$/i.test(name)) {
                SC_alert('alertPickPdfAi');
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
                } catch (_) { }

                // 既存ロジックをそのまま使って総ページ数を取得
                var lastFromFile = getLastPageFromSelection([placedTemp]);
                if (setPathTextFn) setPathTextFn(fileObjOrNull);
                return lastFromFile;
            } finally {
                if (placedTemp) {
                    try { placedTemp.remove(); } catch (_) { }
                }
            }
        }

        // ファイル未指定時は現在の選択から取得（PlacedItem が選択されている前提）
        var last = getLastPageFromSelection(doc.selection);

        // 選択中の配置画像があれば、そのファイルパス表示も更新
        try {
            var placedSel = pageCountFindFirstPlacedItem(doc.selection);
            if (placedSel && placedSel.file && setPathTextFn) setPathTextFn(placedSel.file);
        } catch (_) { }

        return last;
    } catch (e) {
        SC_alert("alertPageCountFail", e);
        return null;
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

// PDF/AIファイルから総ページ数を推定（既存ロジックを維持しつつ簡素化）
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

/**
 * Illustrator のバージョン差異を吸収するため、複数キーに試行します。
 * 期待する値（多くの環境で）: 0=Media, 1=Crop, 2=Bleed, 3=Trim, 4=Art
 */

function isPdfLikeFile(f) {
    var n = String((f && f.name) || "").toLowerCase();
    return (n.indexOf(".pdf") > -1) || (n.indexOf(".ai") > -1);
}


// 入力された文字列（例: "1-20", "1,3,5"）をページ番号配列に変換
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
    var doc = srcDoc;
    // 現在の読み込み対象ファイル。未選択時は null。
    var sourceFile = null;
    var detectedRangeText = "";
    var customRangeText = "";

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

        var pnlAB = leftCol.add('panel', undefined, L('panelLoad'));
        pnlAB.orientation = 'column';
        pnlAB.alignChildren = ['left', 'top'];
        pnlAB.margins = [15, 20, 15, 10];
        var rowRangeMode = pnlAB.add('group');
        rowRangeMode.orientation = 'column';
        rowRangeMode.alignChildren = ['left', 'top'];
        var rbRangeAll = rowRangeMode.add('radiobutton', undefined, L('rangeAll'));
        var rbRangeFirst = rowRangeMode.add('radiobutton', undefined, L('rangeFirst'));
        var rbRangeCustom = rowRangeMode.add('radiobutton', undefined, L('rangeCustom'));
        // 初期状態は「全ページ」を選択
        rbRangeAll.value = true;

        var rowRange = pnlAB.add('group');
        rowRange.orientation = 'row';
        rowRange.alignChildren = ['left', 'center'];
        var etRange = rowRange.add('edittext', undefined, '');
        etRange.characters = 10;
        etRange.enabled = true;

        var groupTotalPages = pnlAB.add('group');
        groupTotalPages.orientation = 'row';
        groupTotalPages.alignChildren = ['right', 'center'];
        groupTotalPages.alignment = ['fill', 'top'];
        groupTotalPages.add('statictext', undefined, '');
        var stTotalPages = groupTotalPages.add('statictext', undefined, '');
        stTotalPages.characters = 8;
        stTotalPages.justify = 'right';

        var panelMode = leftCol.add("panel", undefined, L("panelMode"));
        panelMode.alignChildren = "left";
        panelMode.margins = [15, 20, 15, 10];
        var rowPlaceMode = panelMode.add("group");
        rowPlaceMode.orientation = "column";
        rowPlaceMode.alignChildren = ["left", "top"];
        var rbPerArtboard = rowPlaceMode.add("radiobutton", undefined, L("placePerArtboard"));
        var rbIgnoreArtboard = rowPlaceMode.add("radiobutton", undefined, L("placeIgnoreArtboard"));
        rbPerArtboard.value = true;

        var panelCrop = rightCol.add("panel", undefined, L("panelItem"));
        panelCrop.alignChildren = "left";
        panelCrop.margins = [15, 20, 15, 10];

        var rowCols = panelCrop.add("group");
        rowCols.orientation = "row";
        rowCols.alignChildren = ["left", "center"];
        var stColsLabel = rowCols.add("statictext", undefined, L("colsLabel"));
        var etCols = rowCols.add("edittext", undefined, L("colsAuto"));
        etCols.characters = 5;
        var rowEstimate = panelCrop.add("group");
        rowEstimate.orientation = "row";
        rowEstimate.alignChildren = ["left", "center"];
        rowEstimate.margins = [20, 0, 0, 0];
        var stEstimateLayout = rowEstimate.add("statictext", undefined, "");
        stEstimateLayout.characters = 10;

        var rowArtboardGap = panelCrop.add("group");
        rowArtboardGap.orientation = "row";
        rowArtboardGap.alignChildren = ["left", "center"];
        var stGapLabel = rowArtboardGap.add("statictext", undefined, L("gapLabel"));
        var etArtboardGap = rowArtboardGap.add("edittext", undefined, String(DEFAULT_ARTBOARD_GAP));
        etArtboardGap.characters = 5;
        var stGapUnit = rowArtboardGap.add("statictext", undefined, L("artboardGapUnit"));

        var rowScale = panelCrop.add("group");
        rowScale.orientation = "row";
        rowScale.alignChildren = ["left", "center"];
        var stScaleLabel = rowScale.add("statictext", undefined, L("scaleLabel"));
        var etScale = rowScale.add("edittext", undefined, "100");
        etScale.characters = 5;
        var stScaleUnit = rowScale.add("statictext", undefined, L("scaleUnit"));
        var ddCrop = panelCrop.add("dropdownlist", undefined, [L("cropArt"), L("cropTrim"), L("cropCrop"), L("cropBleed")]);
        ddCrop.minimumSize.width = 160;
        ddCrop.selection = 2;
        ddCrop.enabled = false;

        var panelKei = rightCol.add("panel", undefined, L("panelKei"));
        panelKei.alignChildren = "left";
        panelKei.margins = [15, 20, 15, 10];
        var rowKeiMode = panelKei.add("group");
        rowKeiMode.orientation = "column";
        rowKeiMode.alignChildren = ["left", "top"];
        var rbKeiNone = rowKeiMode.add("radiobutton", undefined, L("keiNone"));
        var rbKeiClipGroup = rowKeiMode.add("radiobutton", undefined, L("keiClipGroup"));
        rbKeiNone.value = true;

        var rowRoundCorner = panelKei.add("group");
        rowRoundCorner.orientation = "row";
        rowRoundCorner.alignChildren = ["left", "center"];
        var cbRoundCorner = rowRoundCorner.add("checkbox", undefined, L("keiRoundCorner"));
        var etRoundCorner = rowRoundCorner.add("edittext", undefined, "3");
        etRoundCorner.characters = 5;
        var stRoundCornerUnit = rowRoundCorner.add("statictext", undefined, getCurrentUnitLabel());
        cbRoundCorner.value = false;
        etRoundCorner.enabled = false;

        var progressBar = win.add("progressbar", undefined, 0, 100);
        progressBar.preferredSize.width = 300;
        progressBar.preferredSize.height = 8;
        progressBar.visible = false;

        var groupButtons = win.add("group");
        groupButtons.orientation = "row";
        groupButtons.alignChildren = ["center", "center"];
        var btnCancel = groupButtons.add("button", undefined, L("cancel"), { name: "cancel" });
        var btnOk = groupButtons.add("button", undefined, L("ok"), { name: "ok" });

        return {
            win: win,
            btnBrowse: btnBrowse,
            etPath: etPath,
            stTotalPages: stTotalPages,
            pnlAB: pnlAB,
            panelMode: panelMode,
            panelCrop: panelCrop,
            panelKei: panelKei,
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
    var etPath = ui.etPath;
    var stTotalPages = ui.stTotalPages;
    var pnlAB = ui.pnlAB;
    var panelMode = ui.panelMode;
    var panelCrop = ui.panelCrop;
    var panelKei = ui.panelKei;
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
        pnlAB.enabled = enabled;
        panelMode.enabled = enabled;
        panelCrop.enabled = enabled;
        panelKei.enabled = enabled;
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
            placedPages = parsePageNumbers(customRangeText || detectedRangeText || '').length;
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
            placedPages = parsePageNumbers(customRangeText || detectedRangeText || '').length;
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
                name: L('notSelected'),
                path: ''
            };
        }

        var nameText;
        var pathText;

        try {
            nameText = decodeURIComponent(f.name);
        } catch (_) {
            nameText = String(f.name || L('notSelected'));
        }

        try {
            pathText = decodeURIComponent(f.fsName);
        } catch (_) {
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
        etPath.text = info.name;
        etPath.helpTip = info.path;
    }

    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var textValue = String(editText.text);
            var value;

            if (textValue === L('colsAuto')) {
                if (event.keyName == "Up") {
                    event.preventDefault();
                    editText.text = 1;
                    try {
                        if (typeof editText.onChange === "function") {
                            editText.onChange();
                        }
                    } catch (_) { }
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
            } catch (_) { }
        });
    }

    // ------------------------
    // UIイベント配線
    // ------------------------
    function bindDialogEvents() {
        btnBrowse.onClick = function () {
            var f = File.openDialog(L('dlgPickFile'), isPdfOrAiFile);
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
                SC_alert('alertNeedFile');
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
                finalPages = parsePageNumbers(etRange.text);
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
        var helpText = rbPerArtboard.value ? L('gapHelpPerArtboard') : L('gapHelpIgnoreArtboard');
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
        var activeIdx = doc.artboards.getActiveArtboardIndex();
        var baseRect = doc.artboards[activeIdx].artboardRect;
        var startX = baseRect[0];
        var nextX = startX;
        var baseTop = baseRect[1];

        var abGap = getArtboardGap();
        var colsPerRow = getColsPerRow();
        var abCount = 0;
        var colCount = 0;
        var rowMaxH = 0; // 現在の行で最も高いページの高さ
        var completed = false;
        showProgress();

        try {
            for (var i = 0; i < targetPages.length; i++) {
                var pageNum = parseInt(targetPages[i], 10);
                if (isNaN(pageNum) || pageNum < 1) pageNum = 1;

                var pageSize = placementMeasurePlacedPageSize(doc, sourceFile, pageNum, cropMode);
                var pageW = pageSize.width;
                var pageH = pageSize.height;

                // 指定列数に達した場合、またはカンバス右端を超える場合、次の行へ折り返す
                var wrapByCol = colsPerRow > 0 && colCount >= colsPerRow;
                var wrapByEdge = colsPerRow === 0 && nextX !== startX && nextX + pageW > CANVAS_RIGHT_EDGE;
                if (wrapByCol || wrapByEdge) {
                    baseTop -= rowMaxH + abGap;
                    nextX = startX;
                    rowMaxH = 0;
                    colCount = 0;
                }

                var result = placementPlaceSinglePage(doc, sourceFile, pageNum, pageW, pageH, nextX, baseTop, activeIdx, abCount, cropMode, abGap);

                if (keiOpts && result.item) {
                    try { keiApplyToPlacedItem(doc, result.item, keiOpts); } catch (_) { }
                }

                nextX = result.nextX;
                abCount = result.abCount;
                colCount++;
                if (pageH > rowMaxH) rowMaxH = pageH;
                updateProgress(i + 1, targetPages.length);
            }
            completed = true;
        } catch (e) {
            SC_alert("alertPlaceError", e);
        } finally {
            try { placementResetImportPageNumber(); } catch (_) { }
            hideProgress();
        }

        if (completed) {
            placementFitAllArtboardsInView(doc);
        }
    }

    function placeIgnoringArtboards(targetPages, cropMode, scale, keiOpts) {
        var activeIdx = doc.artboards.getActiveArtboardIndex();
        var baseRect = doc.artboards[activeIdx].artboardRect;
        var startX = baseRect[0];
        var nextX = startX;
        var baseTop = baseRect[1];
        var s = scale / 100;

        var abGap = getArtboardGap();
        var colsPerRow = getColsPerRow();
        var placedItems = [];
        var colCount = 0;
        var rowMaxH = 0; // 現在の行で最も高いページの高さ（スケール適用後）
        var completed = false;
        showProgress();

        try {
            for (var i = 0; i < targetPages.length; i++) {
                var pageNum = parseInt(targetPages[i], 10);
                if (isNaN(pageNum) || pageNum < 1) pageNum = 1;

                var pageSize = placementMeasurePlacedPageSize(doc, sourceFile, pageNum, cropMode);
                var pageW = pageSize.width * s;
                var pageH = pageSize.height * s;

                // 指定列数に達した場合、またはカンバス右端を超える場合、次の行へ折り返す
                var wrapByCol = colsPerRow > 0 && colCount >= colsPerRow;
                var wrapByEdge = colsPerRow === 0 && nextX !== startX && nextX + pageW > CANVAS_RIGHT_EDGE;
                if (wrapByCol || wrapByEdge) {
                    baseTop -= rowMaxH + abGap;
                    nextX = startX;
                    rowMaxH = 0;
                    colCount = 0;
                }

                var placedItem = placementPlacePage(doc, sourceFile, pageNum, [nextX, baseTop], cropMode, scale);

                if (keiOpts) {
                    try { placedItem = keiApplyToPlacedItem(doc, placedItem, keiOpts); } catch (_) { }
                }

                placedItems.push(placedItem);

                nextX += pageW + abGap;
                colCount++;
                if (pageH > rowMaxH) rowMaxH = pageH;
                updateProgress(i + 1, targetPages.length);
            }
            completed = true;
        } catch (e) {
            SC_alert("alertPlaceError", e);
        } finally {
            try { placementResetImportPageNumber(); } catch (_) { }
            hideProgress();
        }

        if (completed) {
            try {
                doc.selection = null;
            } catch (_) { }
            for (var j = 0; j < placedItems.length; j++) {
                try {
                    placedItems[j].selected = true;
                } catch (_) { }
            }
            placementFitItemsInView(doc, placedItems);
        }
    }


    bindDialogEvents();
    initializeDialogState();

    win.show();

}

main();