#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：
SplitBackgroundForTwo

### 更新日：
20260126

### 概要：
2つのオブジェクト（テキスト、パス、グループなど）を選択して実行すると、各オブジェクトの背面に左右2分割の背景を作成します。
実行時に高さ倍率（%）を指定するダイアログが表示され、閉じる前にプレビューを確認できます（デフォルトは200%）。

左右の背景幅は、2つのオブジェクト間のギャップを基準に計算されます。
［バランス］パネルで「なし／左／右」を選択すると、
- 「なし」：左右均等（中央分割）
- 「左」　：左側のマージンを［幅］で指定（右側は自動計算）
- 「右」　：右側のマージンを［幅］で指定（左側は自動計算）

［幅］はスライダーおよび数値入力で指定でき、選択中の2オブジェクト間ギャップを最大値として自動設定されます。
テキストのサイドベアリング等によるズレを減らすため、
一時的にテキストをアウトライン化して外接矩形を計算し、計算後すぐに一時生成物を削除します。
その上で背景長方形を選択オブジェクトの背面に配置し、ダイアログ内でリアルタイムにプレビュー表示します（元のオブジェクトは変更しません）。

高さ（%）および［幅］入力欄では、↑↓キーで±1、Shift+↑↓で±10（10刻みスナップ）、Option+↑↓で±0.1 の増減が可能です。

### 更新履歴：
- v1.0 (20260124) : 初期バージョン
- v1.1 (20260126) : ［バランス］（なし／左／右）と［幅］指定による左右背景の比率調整に対応。［幅］はオブジェクト間ギャップを最大値として自動計算し、スライダー／数値入力／矢印キー操作に対応
*/

// --- Version / バージョン ---

// --- Dialog UI prefs / ダイアログUI設定 ---
var DIALOG_OFFSET_X = 300;
var DIALOG_OFFSET_Y = 0;
var DIALOG_OPACITY = 0.98;

var SCRIPT_VERSION = "v1.1";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    scriptName: {
        ja: "SplitBackgroundForTwo",
        en: "SplitBackgroundForTwo"
    },
    dialogTitle: {
        ja: "2つのオブジェクトの背景を作成",
        en: "Create Background for Two Objects"
    },
    alertOpenDoc: {
        ja: "ドキュメントを開いてください。",
        en: "Please open a document."
    },
    alertSelectTwoText: {
        ja: "2つのテキストオブジェクトを選択してください。",
        en: "Please select two text objects."
    },
    alertSelectTwoItems: {
        ja: "2つのオブジェクトを選択してください。",
        en: "Please select two objects."
    },
    alertSelectTwoTextFrame: {
        ja: "2つのテキストオブジェクト（TextFrame）を選択してください。",
        en: "Please select two text objects (TextFrame)."
    },
    labelHeight: {
        ja: "高さ：",
        en: "Height:"
    },
    labelPercent: {
        ja: "%",
        en: "%"
    },
    labelPreview: {
        ja: "プレビュー",
        en: "Preview"
    },
    labelOverallFrame: {
        ja: "全体の枠",
        en: "Overall frame"
    },
    labelDivider: {
        ja: "区切り線",
        en: "Divider"
    },
    labelFillLeft: {
        ja: "塗り（左）",
        en: "Fill (Left)"
    },
    labelFillRight: {
        ja: "塗り（右）",
        en: "Fill (Right)"
    },
    panelLine: {
        ja: "オプション",
        en: "Options"
    },
    // --- 固定パネルラベル追加 ---
    panelFixed: {
        ja: "バランス",
        en: "Balance"
    },
    fixedNone: {
        ja: "なし",
        en: "None"
    },
    fixedLeft: {
        ja: "左",
        en: "Left"
    },
    fixedRight: {
        ja: "右",
        en: "Right"
    },
    // --- 幅パネルラベル追加 ---
    panelWidth: {
        ja: "幅",
        en: "Width"
    },
    labelWidth: {
        ja: "幅",
        en: "Width"
    },
    labelPercentSign: {
        ja: "%",
        en: "%"
    },
    // ------------------------
    labelStrokeWidth: {
        ja: "線幅",
        en: "Stroke width"
    },
    labelCornerRadius: {
        ja: "角丸",
        en: "Corner radius"
    },
    labelPt: {
        ja: "pt",
        en: "pt"
    },
    labelOK: {
        ja: "OK",
        en: "OK"
    },
    labelCancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    alertHeightInvalid: {
        ja: "高さ（%）は0より大きい数値を入力してください。",
        en: "Enter a height (%) greater than 0."
    }
};


function L(key) {
    var v = LABELS[key];
    if (!v) return key;
    return (v[lang] !== undefined) ? v[lang] : (v.en !== undefined ? v.en : key);
}

/* 単位ラベル取得ユーティリティ / Unit label utilities */
// 設定キー: rulerType / strokeUnits / text/units / text/asianunits
var UNIT_LABEL_MAP = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};

function getUnitLabel(code, prefKey) {
    // code=5 は Q/H（環境設定キーにより表示が変わる）
    if (code === 5) {
        var hKeys = {
            "text/asianunits": true,
            "rulerType": true,
            "strokeUnits": true
        };
        return hKeys[prefKey] ? "H" : "Q";
    }
    return UNIT_LABEL_MAP[code] || "pt";
}

function getPrefUnitCode(prefKey, fallback) {
    try {
        return app.preferences.getIntegerPreference(prefKey);
    } catch (e) {
        return (fallback !== undefined) ? fallback : 2; // default pt
    }
}

function getCurrentUnitLabelByPrefKey(prefKey) {
    var code = getPrefUnitCode(prefKey, 2);
    return getUnitLabel(code, prefKey);
}

/* pt換算ユーティリティ / Unit conversion (to/from pt) */
// NOTE: Illustrator内部はpt想定。UI入力は各prefKeyの単位で受け取り、内部でptへ変換する。
var UNIT_FACTOR_TO_PT = {
    0: 72,                 // in
    1: 72 / 25.4,          // mm
    2: 1,                  // pt
    3: 12,                 // pica
    4: 72 / 2.54,          // cm
    5: (72 / 25.4) * 0.25, // Q/H : 0.25mm
    6: 1,                  // px (Illustrator: 1px = 1pt @72ppi)
    7: 72,                 // ft/in (use inch as base)
    8: 72 / 0.0254,        // m
    9: 72 * 36,            // yd
    10: 72 * 12            // ft
};

function unitToPt(value, prefKey) {
    var v = Number(value);
    if (isNaN(v)) return NaN;
    var code = getPrefUnitCode(prefKey, 2);
    var f = UNIT_FACTOR_TO_PT[code];
    if (!f) f = 1;
    return v * f;
}

function ptToUnit(ptValue, prefKey) {
    var v = Number(ptValue);
    if (isNaN(v)) return NaN;
    var code = getPrefUnitCode(prefKey, 2);
    var f = UNIT_FACTOR_TO_PT[code];
    if (!f) f = 1;
    return v / f;
}

function formatUnitValue(v) {
    // UI表示用：小数第1位まで（整数に近い場合は整数表示）
    if (isNaN(v)) return "";
    var r = Math.round(v * 10) / 10;
    if (Math.abs(r - Math.round(r)) < 1e-9) return String(Math.round(r));
    return String(r);
}

/* ダイアログ位置をずらす / Shift dialog position */
function shiftDialogPosition(dlg, offsetX, offsetY) {
    if (!dlg) return;
    var prev = dlg.onShow;
    dlg.onShow = function () {
        if (typeof prev === "function") {
            try { prev(); } catch (ePrev) { }
        }
        try {
            var currentX = dlg.location[0];
            var currentY = dlg.location[1];
            dlg.location = [currentX + offsetX, currentY + offsetY];
        } catch (e) { }
    };
}

/* ダイアログ透明度 / Set dialog opacity */
function setDialogOpacity(dlg, opacityValue) {
    if (!dlg) return;
    try { dlg.opacity = opacityValue; } catch (e) { }
}

(function () {
    // ドキュメントが開かれているか確認
    if (app.documents.length === 0) {
        alert(L("alertOpenDoc"));
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    // 2つのオブジェクトが選択されているか確認
    if (!sel || sel.length !== 2) {
        alert(L("alertSelectTwoItems"));
        return;
    }

    // --- Temporary marker (cleanup) / 一時生成物マーカー（後片付け） ---
    var TEMP_NOTE = "__SplitBackgroundForTwo_TEMP__";

    // --- Preview marker (cleanup) / プレビューマーカー（後片付け） ---
    var PREVIEW_NOTE = "__SplitBackgroundForTwo_PREVIEW__";

    function markPreview(item) {
        if (!item) return;
        try { item.note = PREVIEW_NOTE; } catch (e0) { }
    }

    function unmarkPreview(item) {
        if (!item) return;
        try {
            if (item.note === PREVIEW_NOTE) item.note = "";
        } catch (e0) { }
    }

    function removeMarkedPreviewItems(doc) {
        try {
            for (var i = doc.pageItems.length - 1; i >= 0; i--) {
                var it = doc.pageItems[i];
                var isMarked = false;
                try { isMarked = (it.note === PREVIEW_NOTE); } catch (eNote) { isMarked = false; }
                if (isMarked) {
                    unlockAndShow(it);
                    try { it.remove(); } catch (eRm) { }
                }
            }
        } catch (eLoop) { }
    }

    function unlockAndShow(item) {
        if (!item) return;
        try { item.locked = false; } catch (e1) { }
        try { item.hidden = false; } catch (e2) { }
        // Layer / PageItem
        try { item.visible = true; } catch (e3) { }

        // Layer-specific flags (available on Layer)
        try { item.template = false; } catch (e4) { }
        try { item.printable = true; } catch (e5) { }
    }

    function markTemp(item) {
        if (!item) return;
        try { item.note = TEMP_NOTE; } catch (e0) { }

        // Mark children if possible (GroupItem etc.)
        try {
            if (item.pageItems && item.pageItems.length) {
                for (var i = item.pageItems.length - 1; i >= 0; i--) {
                    markTemp(item.pageItems[i]);
                }
            }
        } catch (eChildren) { }
    }

    function removeMarkedTempItems(doc) {
        // Remove any pageItems in the document that are marked (even if created outside temp layer)
        try {
            for (var i = doc.pageItems.length - 1; i >= 0; i--) {
                var it = doc.pageItems[i];
                var isMarked = false;
                try { isMarked = (it.note === TEMP_NOTE); } catch (eNote) { isMarked = false; }
                if (isMarked) {
                    unlockAndShow(it);
                    try { it.remove(); } catch (eRm) { }
                }
            }
        } catch (eLoop) { }
    }

    function removeAllItemsInLayer(layer) {
        // As a fallback, aggressively clear contents of a temp layer
        if (!layer) return;
        unlockAndShow(layer);
        try {
            // Remove pageItems
            for (var i = layer.pageItems.length - 1; i >= 0; i--) {
                var it = layer.pageItems[i];
                unlockAndShow(it);
                try { it.remove(); } catch (e1) { }
            }
        } catch (e2) { }
        try {
            // Remove sublayers
            for (var j = layer.layers.length - 1; j >= 0; j--) {
                var sub = layer.layers[j];
                unlockAndShow(sub);
                try { sub.remove(); } catch (e3) { }
            }
        } catch (e4) { }
    }

    /**
     * テキストを「一時グループ」へ複製→アウトライン化し、その外接矩形を返す。
     * 返り値には後片付け用に一時生成物（複製テキスト／アウトライン）も含める。
     * 元のテキストは変更しない。
     * @param {TextFrame} tf
     * @param {GroupItem} tempGroup - 一時生成物を格納するグループ
     * @returns {{bounds:number[], outline:PageItem|null, duplicateText:TextFrame}}
     */
    function getOutlineBoundsFromText(tf, tempGroup) {
        // 一時レイヤーへ複製（元の階層を汚さない）
        var dup = tf.duplicate(tempGroup, ElementPlacement.PLACEATBEGINNING);

        markTemp(dup);
        unlockAndShow(dup);

        // アウトライン化（結果は一時レイヤー内に生成される想定）
        var outlineItem = null;
        try {
            outlineItem = dup.createOutline();
            if (outlineItem) {
                markTemp(outlineItem);
                unlockAndShow(outlineItem);
                // In some environments outline may be created outside; move it into tempLayer when possible
                try {
                    if (outlineItem.parent !== tempGroup) {
                        outlineItem.move(tempGroup, ElementPlacement.PLACEATBEGINNING);
                    }
                } catch (eMove) { }
            }
        } catch (eOutline) {
            outlineItem = null;
        }

        if (!outlineItem) {
            // Ensure the duplicate itself is marked
            markTemp(dup);
        }

        // bounds 取得対象：基本はアウトライン化結果、取れない場合は複製テキスト
        var target = outlineItem ? outlineItem : dup;
        var b = target.geometricBounds; // [left, top, right, bottom]

        // ここでは削除しない（呼び出し側で安全に削除する）
        return {
            bounds: [b[0], b[1], b[2], b[3]],
            outline: outlineItem,
            duplicateText: dup
        };
    }

    /**
     * 任意オブジェクトの外接矩形（geometricBounds）を取得する。
     * TextFrame の場合はアウトライン化した複製から取得して精度を上げる。
     * それ以外（パス、グループ等）は複製の geometricBounds を使う。
     * @param {PageItem} item
     * @param {GroupItem} tempGroup
     * @returns {{bounds:number[], outline:PageItem|null, duplicateItem:PageItem}}
     */
    function getBoundsInfo(item, tempGroup) {
        if (!item) {
            return { bounds: [0, 0, 0, 0], outline: null, duplicateItem: null };
        }

        if (item.typename === "TextFrame") {
            var info = getOutlineBoundsFromText(item, tempGroup);
            return {
                bounds: info.bounds,
                outline: info.outline,
                duplicateItem: info.duplicateText
            };
        }

        // Non-text: duplicate and read geometricBounds
        var dup = item.duplicate(tempGroup, ElementPlacement.PLACEATBEGINNING);
        markTemp(dup);
        unlockAndShow(dup);

        var b = dup.geometricBounds;
        return {
            bounds: [b[0], b[1], b[2], b[3]],
            outline: null,
            duplicateItem: dup
        };
    }

    /* 左位置を取得 / Get left position */
    function getItemLeft(item) {
        try {
            if (item && item.geometricBounds && item.geometricBounds.length >= 4) {
                return item.geometricBounds[0];
            }
        } catch (e) { }
        try { return item.left; } catch (e2) { }
        return 0;
    }

    // 色の定義 (RGB)
    var colorLeft = new RGBColor();
    colorLeft.red = 220; colorLeft.green = 220; colorLeft.blue = 220; // 薄いグレー

    var colorRight = new RGBColor();
    colorRight.red = 128; colorRight.green = 128; colorRight.blue = 128; // 濃いグレー

    // K100 (CMYK) for frame / 枠線用 K100
    var colorK100 = new CMYKColor();
    colorK100.cyan = 0;
    colorK100.magenta = 0;
    colorK100.yellow = 0;
    colorK100.black = 100;

    // Round Corners Live Effect (for Overall Frame only) / 角丸（ライブエフェクト：全体の枠のみ）
    function createRoundCornersEffectXML(radius) {
        var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius #value# "/></LiveEffect>';
        return xml.replace('#value#', radius);
    }

    function applyRoundCornersIfNeeded(item, radius) {
        if (!item) return;
        var r = Number(radius);
        if (isNaN(r) || r <= 0) return;
        try {
            var xml = createRoundCornersEffectXML(r);
            item.applyEffect(xml);
        } catch (e) { }
    }

    // --- Fill corner rounding (Left: LT/LB, Right: RT/RB) / 塗りの角丸（左=左上/左下、右=右上/右下） ---
    function clearAnchorSelection(pathItem) {
        if (!pathItem || !pathItem.pathPoints) return;
        try {
            var pp = pathItem.pathPoints;
            for (var i = 0; i < pp.length; i++) {
                pp[i].selected = PathPointSelection.NOSELECTION;
            }
        } catch (e) { }
    }

    function selectRightCornersForRectangles(paths) {
        for (var i = 0; i < paths.length; i++) {
            var it = paths[i];
            if (!it || it.typename !== "PathItem" || !it.closed) continue;
            var pp = it.pathPoints;
            if (!pp || pp.length !== 4) continue;

            var sel = 0;
            for (var a = 0; a < 4; a++) if (pp[a].selected === PathPointSelection.ANCHORPOINT) sel++;
            if (sel > 0 && sel < 4) continue; // respect partial anchor selection

            var idx = [0, 1, 2, 3];
            idx.sort(function (x, y) { return pp[x].anchor[0] - pp[y].anchor[0]; });
            var a2 = idx[2], a3 = idx[3];
            var top = (pp[a2].anchor[1] >= pp[a3].anchor[1]) ? a2 : a3;
            var bot = (top === a2) ? a3 : a2;

            for (a = 0; a < 4; a++) pp[a].selected = PathPointSelection.NOSELECTION;
            pp[top].selected = PathPointSelection.ANCHORPOINT;
            pp[bot].selected = PathPointSelection.ANCHORPOINT;
        }
    }

    function applyFillCornerRounding(pathItem, side, radiusPt) {
        if (!pathItem) return;
        var r = Number(radiusPt);
        if (isNaN(r) || r <= 0) return;

        // Use roundAnyCorner algorithm: it rounds only selected corner anchors.
        clearAnchorSelection(pathItem);
        if (side === "left") {
            selectLeftCornersForRectangles([pathItem]);
        } else if (side === "right") {
            selectRightCornersForRectangles([pathItem]);
        } else {
            return;
        }

        try {
            roundAnyCorner([pathItem], { rr: r });
        } catch (e) { }

        // Clean selection on the new rect
        clearAnchorSelection(pathItem);
    }

    // 選択オブジェクトの位置関係を整理（左側にあるものをitem1、右側をitem2とする）
    var item1, item2;
    if (getItemLeft(sel[0]) < getItemLeft(sel[1])) {
        item1 = sel[0];
        item2 = sel[1];
    } else {
        item1 = sel[1];
        item2 = sel[0];
    }

    // --- 背景を置くターゲットレイヤーの決定 ---
    function getLayerIndex(doc, layer) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i] === layer) return i;
        }
        return -1;
    }

    // レイヤーの前後関係をインデックスで推定する
    // Illustratorのlayers配列は基本的に「上（前面）→下（背面）」順で並ぶ想定。
    // よって、インデックスが大きいほど背面寄りとみなす。
    function getBackmostLayer(doc, layerA, layerB) {
        var ia = getLayerIndex(doc, layerA);
        var ib = getLayerIndex(doc, layerB);
        if (ia < 0) return layerB;
        if (ib < 0) return layerA;
        return (ia > ib) ? layerA : layerB;
    }

    var targetLayerForBackground = getBackmostLayer(doc, item1.layer, item2.layer);

    function safeRedraw() {
        try { app.redraw(); } catch (e) { }
    }

    // --- プレビュー状態 ---
    var previewRect1 = null;
    var previewRect2 = null;
    var previewFrame = null;
    var previewDivider = null;
    var previewApplied = false;
    var lastPreviewPercent = null;
    var lastOverallFrame = false;
    var lastDivider = false;
    var lastFillLeft = true;
    var lastFillRight = true;
    var lastStrokeWidth = 1;
    var lastCornerRadius = 0;
    var lastBalanceMode = 'none'; // 'none' | 'left' | 'right'
    var lastWidthPercent = 0;     // Width in pt (from rulerType UI)

    function clearPreviewRects() {
        // 参照が残っている場合はそれを優先して削除
        try {
            if (previewRect1 && previewRect1.isValid) {
                unlockAndShow(previewRect1);
                try { previewRect1.remove(); } catch (e1a) { }
            }
        } catch (e1) { }
        try {
            if (previewRect2 && previewRect2.isValid) {
                unlockAndShow(previewRect2);
                try { previewRect2.remove(); } catch (e2a) { }
            }
        } catch (e2) { }
        try {
            if (previewFrame && previewFrame.isValid) {
                unlockAndShow(previewFrame);
                try { previewFrame.remove(); } catch (e3a) { }
            }
        } catch (e3) { }
        try {
            if (previewDivider && previewDivider.isValid) {
                unlockAndShow(previewDivider);
                try { previewDivider.remove(); } catch (e4a) { }
            }
        } catch (e4) { }

        // 念のため、マーク済みプレビュー図形をドキュメント全体から掃除
        removeMarkedPreviewItems(doc);

        previewRect1 = null;
        previewRect2 = null;
        previewFrame = null;
        previewDivider = null;
        previewApplied = false;
        lastPreviewPercent = null;
        lastDivider = false;
        lastFillLeft = true;
        lastFillRight = true;
        lastStrokeWidth = 1;
        lastCornerRadius = 0;
        safeRedraw();
    }

    /**
     * アウトライン計算→背景長方形を作成（戻り値で生成物を返す）
     * @param {PageItem} leftText
     * @param {PageItem} rightText
     * @param {number} heightRatioValue
     * @param {boolean} addOverallFrame
     * @param {boolean} addDivider
     * @param {boolean} addFillLeft
     * @param {boolean} addFillRight
     * @param {number} strokeWidthValue
     * @param {number} cornerRadius
     * @returns {{rect1:PathItem|null, rect2:PathItem|null, frame:PathItem|null, divider:PathItem|null}}
     */
    function drawBackgroundRects(leftText, rightText, heightRatioValue, addOverallFrame, addDivider, addFillLeft, addFillRight, strokeWidthValue, cornerRadius, balanceMode, widthPercent) {        // --- 一時グループ（アウトライン生成用）---
        var originalActiveLayer = doc.activeLayer;

        // 一時グループを「既存レイヤー」内に作る（レイヤー残留問題を回避）
        var hostLayer = null;
        try { hostLayer = leftText.layer; } catch (eHL) { hostLayer = null; }
        if (!hostLayer) {
            try { hostLayer = doc.activeLayer; } catch (eHL2) { hostLayer = null; }
        }
        if (!hostLayer) {
            hostLayer = doc.layers[0];
        }

        unlockAndShow(hostLayer);
        var tempGroup = hostLayer.groupItems.add();
        tempGroup.name = "__temp_outline_group__";
        markTemp(tempGroup);
        // 見えにくくする（万一残っても視覚ノイズを下げる）
        try { tempGroup.opacity = 0; } catch (eOp) { }

        // 計算結果（この値だけ持ち帰って、アウトラインはすぐ消す）
        var rectTop, rectHeight;
        var rectLeftStart, gapCenter, rectRightEnd;

        try {
            var outline1Info = getBoundsInfo(leftText, tempGroup);
            var outline2Info = getBoundsInfo(rightText, tempGroup);

            var b1 = outline1Info.bounds;
            var b2 = outline2Info.bounds;

            // --- 計算が済んだら一時生成物を即削除（残留対策）---
            unlockAndShow(outline1Info.outline);
            unlockAndShow(outline1Info.duplicateItem);
            unlockAndShow(outline2Info.outline);
            unlockAndShow(outline2Info.duplicateItem);
            try { if (outline1Info.outline && outline1Info.outline.isValid) outline1Info.outline.remove(); } catch (eA) { }
            try { if (outline1Info.duplicateItem && outline1Info.duplicateItem.isValid) outline1Info.duplicateItem.remove(); } catch (eB) { }
            try { if (outline2Info.outline && outline2Info.outline.isValid) outline2Info.outline.remove(); } catch (eC) { }
            try { if (outline2Info.duplicateItem && outline2Info.duplicateItem.isValid) outline2Info.duplicateItem.remove(); } catch (eD) { }
            removeMarkedTempItems(doc);

            // 1. 高さの計算 (共通)
            var topLimit = Math.max(b1[1], b2[1]);
            var bottomLimit = Math.min(b1[3], b2[3]);

            var contentHeight = topLimit - bottomLimit;
            var centerY = topLimit - (contentHeight / 2);

            rectHeight = contentHeight * heightRatioValue;
            rectTop = centerY + (rectHeight / 2);

            // 2. 幅と隙間の計算（アウトライン bounds ベース）
            var item1Left = b1[0];
            var item1Right = b1[2];
            var item2Left = b2[0];
            var item2Right = b2[2];

            var gapDist = item2Left - item1Right;
            if (gapDist < 0) gapDist = 0;

            // --- Balance / margin control ---
            // widthPercent is treated as an absolute margin in pt.
            function clamp(v, mn, mx) {
                if (v < mn) return mn;
                if (v > mx) return mx;
                return v;
            }

            var wPt = Number(widthPercent);
            if (isNaN(wPt) || wPt < 0) wPt = 0;

            var marginLeft = gapDist / 2;
            var marginRight = gapDist / 2;

            if (balanceMode === 'left') {
                // 左の左右マージン = スライダー値
                marginLeft = clamp(wPt, 0, gapDist);
                // 右の左右マージン = gap - 左
                marginRight = gapDist - marginLeft;

            } else if (balanceMode === 'right') {
                // 右の左右マージン = スライダー値
                marginRight = clamp(wPt, 0, gapDist);
                // 左の左右マージン = gap - 右
                marginLeft = gapDist - marginRight;

            } else {
                // none => centered
                marginLeft = gapDist / 2;
                marginRight = gapDist / 2;
            }

            gapCenter = item1Right + marginLeft;
            rectLeftStart = item1Left - marginLeft;
            rectRightEnd = item2Right + marginRight;

        } finally {
            // 一時グループを確実に削除（レイヤー削除ではなくグループ削除で残留を防ぐ）
            try { doc.selection = null; } catch (eSel) { }

            try {
                if (originalActiveLayer && originalActiveLayer.isValid) {
                    doc.activeLayer = originalActiveLayer;
                }
            } catch (eRestore) { }

            // グループ内の残骸を掃除してから削除
            try {
                unlockAndShow(tempGroup);
                for (var gi = tempGroup.pageItems.length - 1; gi >= 0; gi--) {
                    try {
                        unlockAndShow(tempGroup.pageItems[gi]);
                        tempGroup.pageItems[gi].remove();
                    } catch (eRmPI) { }
                }
            } catch (eClearG) { }

            try {
                if (tempGroup && tempGroup.isValid) {
                    tempGroup.remove();
                }
            } catch (eRemoveG) { }

            // 最終保険：マーク付き残骸をもう一度掃除
            removeMarkedTempItems(doc);
        }

        // --- ここから描画（アウトラインは既に削除済み）---
        // 背景長方形は必ずターゲットレイヤー直下に作成
        var leftLayer = null;
        var rightLayer = null;
        try { leftLayer = leftText.layer; } catch (eLL) { leftLayer = null; }
        try { rightLayer = rightText.layer; } catch (eRL) { rightLayer = null; }

        var rect1 = null;
        var rect2 = null;

        if (addFillLeft) {
            rect1 = (leftLayer ? leftLayer.pathItems : doc.pathItems).rectangle(
                rectTop,
                rectLeftStart,
                gapCenter - rectLeftStart,
                rectHeight
            );
            rect1.fillColor = colorLeft;
            rect1.stroked = false;
        }

        if (addFillRight) {
            rect2 = (rightLayer ? rightLayer.pathItems : doc.pathItems).rectangle(
                rectTop,
                gapCenter,
                rectRightEnd - gapCenter,
                rectHeight
            );
            rect2.fillColor = colorRight;
            rect2.stroked = false;
        }

        // 塗りの角丸（左=左上/左下、右=右上/右下） / Fill corner rounding
        if (cornerRadius && Number(cornerRadius) > 0) {
            if (rect1) applyFillCornerRounding(rect1, "left", cornerRadius);
            if (rect2) applyFillCornerRounding(rect2, "right", cornerRadius);
        }

        var frameRect = null;
        if (addOverallFrame) {
            // Create a single stroked rectangle that surrounds both background rectangles
            // 枠線：2つの背景を囲む単一の長方形（strokeWidthValue / K100）
            try {
                var frameLayer = targetLayerForBackground || leftLayer || rightLayer;
                frameRect = (frameLayer ? frameLayer.pathItems : doc.pathItems).rectangle(
                    rectTop,
                    rectLeftStart,
                    rectRightEnd - rectLeftStart,
                    rectHeight
                );
                frameRect.filled = false;
                frameRect.stroked = true;
                frameRect.strokeWidth = strokeWidthValue;
                frameRect.strokeColor = colorK100;

                // 角丸（全体の枠のみ） / Corner radius (Overall Frame only)
                applyRoundCornersIfNeeded(frameRect, cornerRadius);

                // Ensure frame is above the fill rectangles
                try { frameRect.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (eFZ) { }
            } catch (eFrame) {
                frameRect = null;
            }
        }

        var dividerLine = null;
        if (addDivider) {
            // Divider line at the center boundary / 区切り線（中央境界）
            try {
                var dividerLayer = targetLayerForBackground || leftLayer || rightLayer;
                dividerLine = (dividerLayer ? dividerLayer.pathItems : doc.pathItems).add();
                dividerLine.setEntirePath([
                    [gapCenter, rectTop],
                    [gapCenter, rectTop - rectHeight]
                ]);
                dividerLine.filled = false;
                dividerLine.stroked = true;
                dividerLine.strokeWidth = strokeWidthValue;
                dividerLine.strokeColor = colorK100;
                try { dividerLine.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (eDZ) { }
            } catch (eDiv) {
                dividerLine = null;
            }
        }

        // 背面へ（レイヤー内）
        try { if (rect1) rect1.zOrder(ZOrderMethod.SENDTOBACK); } catch (eZ1) { }
        try { if (rect2) rect2.zOrder(ZOrderMethod.SENDTOBACK); } catch (eZ2) { }

        // テキストを前面へ（同一レイヤー/同一親内での保証）
        try { leftText.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (eT1) { }
        try { rightText.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (eT2) { }

        safeRedraw();
        return { rect1: rect1, rect2: rect2, frame: frameRect, divider: dividerLine };
    }

    /**
     * ↑↓キーで数値を増減するユーティリティ
     * - ↑↓ : ±1
     * - Shift + ↑↓ : ±10（10刻みにスナップ）
     * - Option(Alt) + ↑↓ : ±0.1
     * @param {EditText} editText
     * @param {boolean} allowNegative
     * @param {Function} onChange - 値変更後に呼ぶ（プレビュー更新など）
     */
    function changeValueByArrowKey(editText, allowNegative, onChange) {
        if (!editText) return;

        editText.addEventListener("keydown", function (event) {
            // Up/Down 以外は何もしない
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                } else if (event.keyName === "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                }
                event.preventDefault();
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Option(Alt)キー押下時は0.1単位で増減
                if (event.keyName === "Up") {
                    value += delta;
                } else if (event.keyName === "Down") {
                    value -= delta;
                }
                event.preventDefault();
            } else {
                delta = 1;
                if (event.keyName === "Up") {
                    value += delta;
                } else if (event.keyName === "Down") {
                    value -= delta;
                }
                event.preventDefault();
            }

            // 丸め
            if (keyboard.altKey) {
                value = Math.round(value * 10) / 10; // 小数第1位
            } else {
                value = Math.round(value); // 整数
            }

            if (!allowNegative && value < 0) value = 0;

            editText.text = String(value);

            // 値変更後の処理（プレビュー更新など）
            if (typeof onChange === "function") {
                try { onChange(); } catch (e) { }
            }
        });
    }

    // --- ダイアログ：高さ（%）指定（プレビュー付き）/ Dialog: Height (%) (with preview) ---
    /**
     * 高さ（%）を指定するダイアログ（プレビュー付き）
     * @param {number} defaultPercent
     * @param {Function} previewFn - function(percent:number, enabled:boolean, addOverallFrame:boolean, addDivider:boolean, addFillLeft:boolean, addFillRight:boolean, strokeWidth:number, cornerRadius:number):void
     * @param {Function} clearPreviewFn - function():void
     * @returns {number|null} percent (e.g. 140) or null if cancelled
     */
    function showHeightDialog(defaultPercent, previewFn, clearPreviewFn) {
        var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
        dlg.orientation = 'column';
        dlg.alignChildren = ['fill', 'top'];
        dlg.margins = 18;

        // ダイアログ透明度・位置調整 / Dialog opacity & position
        setDialogOpacity(dlg, DIALOG_OPACITY);
        shiftDialogPosition(dlg, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);

        var row = dlg.add('group');
        row.orientation = 'row';
        row.alignChildren = ['left', 'center'];
        // 高さ行は左右中央に / Center the height row horizontally
        row.alignment = ['center', 'top'];

        row.add('statictext', undefined, L('labelHeight'));

        var et = row.add('edittext', undefined, String(defaultPercent));
        et.characters = 6;
        et.active = true;
        // ↑↓ / Shift+↑↓ / Option+↑↓ で値を増減
        changeValueByArrowKey(et, false, applyPreview);

        row.add('statictext', undefined, L('labelPercent'));

        // --- 2カラム：左=描画 / 右=オプション ---
        var columns = dlg.add('group');
        columns.orientation = 'row';
        columns.alignChildren = ['fill', 'top'];
        columns.spacing = 12;

        // --- 描画オプション / Drawing options (Left column) ---
        var drawPanel = columns.add('panel', undefined, '描画');
        drawPanel.orientation = 'column';
        drawPanel.alignChildren = ['left', 'top'];
        drawPanel.margins = [15, 20, 15, 10];

        var cbFillLeft = drawPanel.add('checkbox', undefined, L('labelFillLeft'));
        cbFillLeft.value = true;

        var cbFillRight = drawPanel.add('checkbox', undefined, L('labelFillRight'));
        cbFillRight.value = true;


        var cbOverallFrame = drawPanel.add('checkbox', undefined, L('labelOverallFrame'));
        cbOverallFrame.value = false;

        var cbDivider = drawPanel.add('checkbox', undefined, L('labelDivider'));
        cbDivider.value = false;

        // --- 線オプション / Stroke options (Right column) ---
        var rightCol = columns.add('group');
        rightCol.orientation = 'column';
        rightCol.alignChildren = ['fill', 'top'];
        rightCol.spacing = 12;

        var linePanel = rightCol.add('panel', undefined, L('panelLine'));
        linePanel.orientation = 'column';
        linePanel.alignChildren = ['left', 'top'];
        linePanel.margins = [15, 20, 15, 10];


        var lineRow = linePanel.add('group');
        lineRow.orientation = 'row';
        lineRow.alignChildren = ['left', 'center'];

        lineRow.add('statictext', undefined, L('labelStrokeWidth'));
        var etStroke = lineRow.add('edittext', undefined, formatUnitValue(ptToUnit(1, "strokeUnits")));
        etStroke.characters = 4;
        // Allow decimals and arrow-key changes (0.1 with Option)
        changeValueByArrowKey(etStroke, false, applyPreview);
        lineRow.add('statictext', undefined, getCurrentUnitLabelByPrefKey("strokeUnits"));

        // 角丸 / Corner radius（枠のみロジック対応。塗りは追って）
        var cornerRow = linePanel.add('group');
        cornerRow.orientation = 'row';
        cornerRow.alignChildren = ['left', 'center'];

        cornerRow.add('statictext', undefined, L('labelCornerRadius'));
        var etCorner = cornerRow.add('edittext', undefined, formatUnitValue(ptToUnit(0, "rulerType")));
        etCorner.characters = 4;
        cornerRow.add('statictext', undefined, getCurrentUnitLabelByPrefKey("rulerType"));
        // ↑↓ / Shift+↑↓ / Option+↑↓ で角丸値を増減
        changeValueByArrowKey(etCorner, false, applyPreview);

        // --- Helpers: enable/disable UI rows ---
        function updateStrokeWidthEnabled() {
            var on = !!(cbOverallFrame.value || cbDivider.value);
            // 両方OFFのときは線幅UIをディム表示 / Dim stroke width UI when both are OFF
            try { lineRow.enabled = on; } catch (eEn1) { }
        }

        function updateCornerEnabled() {
            var on = !!((cbFillLeft.value || cbFillRight.value) || cbOverallFrame.value);
            // 両方OFFのときは角丸UIをディム表示 / Dim corner radius UI when both are OFF
            try { cornerRow.enabled = on; } catch (eEn2) { }
        }

        updateStrokeWidthEnabled();
        updateCornerEnabled();

        // --- バランス / Balance (full-width, spanning both columns) ---
        var pinPanel = dlg.add('panel', undefined, L('panelFixed'));
        pinPanel.orientation = 'column';
        pinPanel.alignChildren = ['fill', 'top'];
        pinPanel.margins = [15, 20, 15, 10];
        pinPanel.alignment = ['fill', 'top'];

        var pinRadioRow = pinPanel.add('group');
        pinRadioRow.orientation = 'row';
        pinRadioRow.alignChildren = ['left', 'center'];

        var rbPinNone = pinRadioRow.add('radiobutton', undefined, L('fixedNone'));
        var rbPinLeft = pinRadioRow.add('radiobutton', undefined, L('fixedLeft'));
        var rbPinRight = pinRadioRow.add('radiobutton', undefined, L('fixedRight'));
        rbPinNone.value = true;

        // --- 幅 / Width (inside Balance panel) ---
        var pinWidthCol = pinPanel.add('group');
        pinWidthCol.orientation = 'column';
        pinWidthCol.alignChildren = ['fill', 'top'];
        pinWidthCol.spacing = 6;

        // Value row : 幅 [ value ] unit
        var pinWidthValueRow = pinWidthCol.add('group');
        pinWidthValueRow.orientation = 'row';
        pinWidthValueRow.alignChildren = ['left', 'center'];

        pinWidthValueRow.add('statictext', undefined, L('labelWidth'));

        var etWidth = pinWidthValueRow.add('edittext', undefined, '0');
        etWidth.characters = 6;
        // ↑↓ / Shift+↑↓ / Option+↑↓ で幅を増減（プレビュー更新）
        changeValueByArrowKey(etWidth, false, applyPreview);

        pinWidthValueRow.add('statictext', undefined, getCurrentUnitLabelByPrefKey("rulerType"));

        // Slider row
        var pinWidthSliderRow = pinWidthCol.add('group');
        pinWidthSliderRow.orientation = 'row';
        pinWidthSliderRow.alignChildren = ['fill', 'center'];

        // Slider: 0 - max (will be set dynamically)
        var slWidth = pinWidthSliderRow.add('slider', undefined, 0, 0, 0); // max will be set dynamically
        slWidth.preferredSize = [220, 20];

        // Keep slider and edittext in sync (UI only)
        function clampWidthValue(v, forceInteger) {
            v = Number(v);
            if (isNaN(v)) return 0;
            if (v < 0) v = 0;

            var maxV = 0;
            try { maxV = Number(slWidth.maxvalue); } catch (eMax0) { maxV = 0; }
            if (isNaN(maxV) || maxV < 0) maxV = 0;

            if (v > maxV) v = maxV;

            if (forceInteger) {
                return Math.round(v); // integer
            }
            return Math.round(v * 10) / 10; // keep 0.1 precision
        }

        function syncWidthUIFromSlider() {
            var kb = ScriptUI.environment.keyboardState;
            var forceInt = !!(kb && kb.altKey);
            var v = clampWidthValue(slWidth.value, forceInt);
            slWidth.value = v;
            etWidth.text = String(v);
        }

        function syncWidthUIFromEdit() {
            var v = clampWidthValue(etWidth.text, false);
            slWidth.value = v;
            etWidth.text = String(v);
        }

        // --- Compute and set the max value for the width slider based on selection ---
        function updateWidthSliderMaxBySelection() {
            // Width max should be the gap distance between the two objects (in rulerType units)
            // Use outline bounds for TextFrame to avoid side-bearing issues (same as main logic)
            var maxUnit = 0;
            var tempGroupForMax = null;
            try {
                var hostLayer2 = null;
                try { hostLayer2 = item1.layer; } catch (eHLm) { hostLayer2 = null; }
                if (!hostLayer2) {
                    try { hostLayer2 = doc.activeLayer; } catch (eHLm2) { hostLayer2 = null; }
                }
                if (!hostLayer2) hostLayer2 = doc.layers[0];

                unlockAndShow(hostLayer2);
                tempGroupForMax = hostLayer2.groupItems.add();
                tempGroupForMax.name = "__temp_outline_group_for_max__";
                markTemp(tempGroupForMax);
                try { tempGroupForMax.opacity = 0; } catch (eOpM) { }

                var bi1 = getBoundsInfo(item1, tempGroupForMax);
                var bi2 = getBoundsInfo(item2, tempGroupForMax);
                var b1 = bi1.bounds;
                var b2 = bi2.bounds;

                // ensure left/right by x
                var item1Right = b1[2];
                var item2Left = b2[0];
                var gapPt = item2Left - item1Right;
                if (gapPt < 0) gapPt = 0;

                maxUnit = ptToUnit(gapPt, "rulerType");
                if (isNaN(maxUnit) || maxUnit < 0) maxUnit = 0;
                // keep 0.1 in UI
                maxUnit = Math.round(maxUnit * 10) / 10;

            } catch (eMax) {
                maxUnit = 0;
            } finally {
                // cleanup temp group
                try { removeMarkedTempItems(doc); } catch (eRmTmp0) { }
                try {
                    if (tempGroupForMax && tempGroupForMax.isValid) {
                        unlockAndShow(tempGroupForMax);
                        for (var giM = tempGroupForMax.pageItems.length - 1; giM >= 0; giM--) {
                            try {
                                unlockAndShow(tempGroupForMax.pageItems[giM]);
                                tempGroupForMax.pageItems[giM].remove();
                            } catch (eRmPI2) { }
                        }
                        tempGroupForMax.remove();
                    }
                } catch (eRmTmp1) { }
                try { removeMarkedTempItems(doc); } catch (eRmTmp2) { }
            }

            try { slWidth.maxvalue = maxUnit; } catch (eSetMax) { }
            // If max becomes 0, keep slider usable but effectively fixed at 0
            try { if (slWidth.maxvalue < 0) slWidth.maxvalue = 0; } catch (eSetMax2) { }

            // Clamp current UI value to new max
            syncWidthUIFromEdit();

            // When "None" is selected, row is disabled anyway; this is just the upper bound.
        }

        slWidth.onChanging = function () {
            syncWidthUIFromSlider();
            applyPreview(); // NOTE: logic will be wired later; keep preview refresh for now
        };

        etWidth.addEventListener('changing', function () {
            syncWidthUIFromEdit();
            applyPreview();
        });

        // --- Helpers: enable/disable width slider by balance selection ---
        function updateWidthEnabled() {
            // Disable width slider when "None" is selected
            try { pinWidthCol.enabled = !rbPinNone.value; } catch (e) { }
        }

        rbPinNone.onClick = function () { updateWidthEnabled(); applyPreview(); };
        rbPinLeft.onClick = function () { updateWidthEnabled(); applyPreview(); };
        rbPinRight.onClick = function () { updateWidthEnabled(); applyPreview(); };

        // Initial state
        updateWidthEnabled();


        // プレビューは一番下 / Preview at the bottom
        var cbPreview = dlg.add('checkbox', undefined, L('labelPreview'));
        cbPreview.value = true;

        function parsePercent() {
            var v = Number(et.text);
            if (isNaN(v) || v <= 0) return null;
            return v;
        }

        function parseStrokeWidth() {
            var vUnit = Number(etStroke.text);
            if (isNaN(vUnit) || vUnit <= 0) return null;
            // round to 0.1 (UI単位)
            vUnit = Math.round(vUnit * 10) / 10;
            var vPt = unitToPt(vUnit, "strokeUnits");
            if (isNaN(vPt) || vPt <= 0) return null;
            // internal pt: keep 0.1pt precision
            vPt = Math.round(vPt * 10) / 10;
            return vPt;
        }

        function parseCornerRadius() {
            var vUnit = Number(etCorner.text);
            if (isNaN(vUnit) || vUnit < 0) return null;
            // round to 0.1 (UI単位)
            vUnit = Math.round(vUnit * 10) / 10;
            var vPt = unitToPt(vUnit, "rulerType");
            if (isNaN(vPt) || vPt < 0) return null;
            // internal pt: keep 0.1pt precision
            vPt = Math.round(vPt * 10) / 10;
            return vPt;
        }

        function getBalanceMode() {
            if (rbPinLeft.value) return 'left';
            if (rbPinRight.value) return 'right';
            return 'none';
        }

        function parseWidthPercent() {
            // NOTE: 「幅」は距離（rulerType単位）として扱い、内部はptに変換する
            var vUnit = Number(etWidth.text);
            if (isNaN(vUnit) || vUnit < 0) vUnit = 0;

            // keep 0.1 in UI units
            vUnit = Math.round(vUnit * 10) / 10;

            // keep UI in sync
            try { slWidth.value = vUnit; } catch (e0) { }
            try { etWidth.text = String(vUnit); } catch (e1) { }

            var vPt = unitToPt(vUnit, "rulerType");
            if (isNaN(vPt) || vPt < 0) vPt = 0;

            // internal pt: keep 0.1pt precision
            vPt = Math.round(vPt * 10) / 10;
            return vPt;
        }

        function applyPreview() {
            var v = parsePercent();
            if (v === null) {
                // 数値として成立しない間は、プレビューを消すだけ
                clearPreviewFn();
                return;
            }
            var sw = parseStrokeWidth();
            if (sw === null) {
                clearPreviewFn();
                return;
            }
            var cr = parseCornerRadius();
            if (cr === null) {
                clearPreviewFn();
                return;
            }
            var bm = getBalanceMode();
            var wp = parseWidthPercent();
            previewFn(v, cbPreview.value, cbOverallFrame.value, cbDivider.value, cbFillLeft.value, cbFillRight.value, sw, cr, bm, wp);
        }

        // ダイアログ表示直後に初回プレビューを実行 / Run initial preview on dialog show
        (function () {
            var prevOnShow = dlg.onShow;
            dlg.onShow = function () {
                if (typeof prevOnShow === "function") {
                    try { prevOnShow(); } catch (ePrevShow) { }
                }
                try { updateWidthSliderMaxBySelection(); } catch (eMaxInit) { }
                try { applyPreview(); } catch (eInitPrev) { }
            };
        })();

        // 値変更で即プレビュー
        et.addEventListener('changing', function () {
            applyPreview();
        });
        etStroke.addEventListener('changing', function () {
            applyPreview();
        });
        etCorner.addEventListener('changing', function () {
            applyPreview();
        });

        cbPreview.onClick = function () {
            applyPreview();
        };
        cbOverallFrame.onClick = function () {
            updateStrokeWidthEnabled();
            updateCornerEnabled();
            applyPreview();
        };
        cbDivider.onClick = function () {
            updateStrokeWidthEnabled();
            applyPreview();
        };
        cbFillLeft.onClick = function () {
            updateCornerEnabled();
            applyPreview();
        };
        cbFillRight.onClick = function () {
            updateCornerEnabled();
            applyPreview();
        };

        // OK/Cancel
        var btns = dlg.add('group');
        btns.orientation = 'row';
        btns.alignment = ['right', 'center'];

        var cancelBtn = btns.add('button', undefined, L('labelCancel'), { name: 'cancel' });
        var okBtn = btns.add('button', undefined, L('labelOK'), { name: 'ok' });

        okBtn.onClick = function () {
            var v = parsePercent();
            if (v === null) {
                alert(L('alertHeightInvalid'));
                return;
            }
            // OK：値だけ確定（プレビューは既に反映済み）
            // プレビューOFFの場合はここで消してから確定
            if (!cbPreview.value) {
                clearPreviewFn();
            }
            dlg.close(1);
        };

        cancelBtn.onClick = function () {
            clearPreviewFn();
            dlg.close(0);
        };

        // Enter / Esc
        dlg.defaultElement = okBtn;
        dlg.cancelElement = cancelBtn;

        var result = dlg.show();
        if (result !== 1) return null;

        var percent = parsePercent();
        if (percent === null) return null;

        return percent;
    }

    var DEFAULT_HEIGHT_PERCENT = 200; // ダイアログ初期値（%）
    var heightRatio = DEFAULT_HEIGHT_PERCENT / 100; // 高さ倍率（内部計算用）

    function previewFn(percent, enabled, addOverallFrame, addDivider, addFillLeft, addFillRight, strokeWidth, cornerRadius, balanceMode, widthPercent) {
        clearPreviewRects();
        if (!enabled) {
            return;
        }
        try {
            var hr = percent / 100;
            var result = drawBackgroundRects(item1, item2, hr, addOverallFrame, addDivider, addFillLeft, addFillRight, strokeWidth, cornerRadius, balanceMode, widthPercent);
            previewRect1 = result.rect1;
            previewRect2 = result.rect2;
            previewFrame = result.frame;
            previewDivider = result.divider;
            // プレビュー生成物としてマーク（参照が失われても掃除できるように）
            if (previewRect1) markPreview(previewRect1);
            if (previewRect2) markPreview(previewRect2);
            if (previewFrame) markPreview(previewFrame);
            if (previewDivider) markPreview(previewDivider);
            previewApplied = true;
            lastPreviewPercent = percent;
            lastOverallFrame = !!addOverallFrame;
            lastDivider = !!addDivider;
            lastFillLeft = !!addFillLeft;
            lastFillRight = !!addFillRight;
            lastStrokeWidth = strokeWidth;
            lastCornerRadius = cornerRadius;
            lastBalanceMode = balanceMode || 'none';
            lastWidthPercent = (widthPercent !== undefined && widthPercent !== null) ? widthPercent : 0;
            safeRedraw();
        } catch (ePrev) {
            // 途中生成が残らないように
            clearPreviewRects();
            throw ePrev;
        }
    }

    var heightPercent = showHeightDialog(DEFAULT_HEIGHT_PERCENT, previewFn, clearPreviewRects);
    if (heightPercent === null) {
        return; // キャンセル
    }
    heightRatio = heightPercent / 100;

    // --- 確定処理 ---
    // プレビューが有効で既に描画済みなら、それをそのまま採用する
    if (!(previewApplied && lastPreviewPercent === heightPercent)) {
        // 既存プレビューがない、または別の値のプレビューだった場合はここで描画
        var finalResult = drawBackgroundRects(item1, item2, heightRatio, lastOverallFrame, lastDivider, lastFillLeft, lastFillRight, lastStrokeWidth, lastCornerRadius, lastBalanceMode, lastWidthPercent); previewRect1 = finalResult.rect1;
        previewRect2 = finalResult.rect2;
        previewFrame = finalResult.frame;
        previewDivider = finalResult.divider;
    } else {
        // プレビュー図形をそのまま採用するので、プレビュー用マークを外す
        if (previewRect1) unmarkPreview(previewRect1);
        if (previewRect2) unmarkPreview(previewRect2);
        unmarkPreview(previewFrame);
        unmarkPreview(previewDivider);
    }

    // 選択を解除
    doc.selection = null;

    // ==============================
    // Round Any Corner / 角丸アルゴリズム（選択アンカーのみ丸める）
    // Based on: Hiroyuki Sato (MIT) https://github.com/shspage
    // ==============================

    function roundAnyCorner(s, conf) {
        var rr = conf.rr;

        // var tim = new Date();
        var p, op, pnts;
        var skipList, adjRdirAtEnd, redrawFlg;
        var i, nxi, pvi, q, d, ds, r, g, t, qb;
        var anc1, ldir1, rdir1, anc2, ldir2, rdir2;

        var hanLen = 4 * (Math.sqrt(2) - 1) / 3;
        var ptyp = PointType.SMOOTH;

        for (var j = 0; j < s.length; j++) {
            p = s[j].pathPoints;
            if (readjustAnchors(p) < 2) continue; // reduce anchors
            op = !s[j].closed;
            pnts = op ? [getDat(p[0])] : [];
            redrawFlg = false;
            adjRdirAtEnd = 0;

            skipList = [(op || !isSelected(p[0]) || !isCorner(p, 0))];
            for (i = 1; i < p.length; i++) {
                skipList.push((!isSelected(p[i])
                    || !isCorner(p, i)
                    || (op && i == p.length - 1)));
            }

            for (i = 0; i < p.length; i++) {
                nxi = parseIdx(p, i + 1);
                if (nxi < 0) break;

                pvi = parseIdx(p, i - 1);

                q = [p[i].anchor, p[i].rightDirection,
                p[nxi].leftDirection, p[nxi].anchor];

                ds = dist(q[0], q[3]) / 2;
                if (arrEq(q[0], q[1]) && arrEq(q[2], q[3])) {  // straight side
                    r = Math.min(ds, rr);
                    g = getRad(q[0], q[3]);
                    anc1 = getPnt(q[0], g, r);
                    ldir1 = getPnt(anc1, g + Math.PI, r * hanLen);

                    if (skipList[nxi]) {
                        if (!skipList[i]) {
                            pnts.push([anc1, anc1, ldir1, ptyp]);
                            redrawFlg = true;
                        }
                        pnts.push(getDat(p[nxi]));
                    } else {
                        if (r < rr) {  // when the length of the side is less than rr * 2
                            pnts.push([anc1,
                                getPnt(anc1, getRad(ldir1, anc1), r * hanLen),
                                ldir1,
                                ptyp]);
                        } else {
                            if (!skipList[i]) pnts.push([anc1, anc1, ldir1, ptyp]);
                            anc2 = getPnt(q[3], g + Math.PI, r);
                            pnts.push([anc2,
                                getPnt(anc2, g, r * hanLen),
                                anc2,
                                ptyp]);
                        }
                        redrawFlg = true;
                    }
                } else {  // not straight side
                    d = getT4Len(q, 0) / 2;
                    r = Math.min(d, rr);
                    t = getT4Len(q, r);
                    anc1 = bezier(q, t);
                    rdir1 = defHan(t, q, 1);
                    ldir1 = getPnt(anc1, getRad(rdir1, anc1), r * hanLen);

                    if (skipList[nxi]) {
                        if (skipList[i]) {
                            pnts.push(getDat(p[nxi]));
                        } else {
                            pnts.push([anc1, rdir1, ldir1, ptyp]);
                            with (p[nxi]) pnts.push([anchor,
                                rightDirection,
                                adjHan(anchor, leftDirection, 1 - t),
                                ptyp]);
                            redrawFlg = true;
                        }
                    } else { // skipList[nxi] = false
                        if (r < rr) {  // the length of the side is less than rr * 2
                            if (skipList[i]) {
                                if (!op && i == 0) {
                                    adjRdirAtEnd = t;
                                } else {
                                    pnts[pnts.length - 1][1] = adjHan(q[0], q[1], t);
                                }
                                pnts.push([anc1,
                                    getPnt(anc1, getRad(ldir1, anc1), r * hanLen),
                                    defHan(t, q, 0),
                                    ptyp]);
                            } else {
                                pnts.push([anc1,
                                    getPnt(anc1, getRad(ldir1, anc1), r * hanLen),
                                    ldir1,
                                    ptyp]);
                            }
                        } else {  // round the corner with the radius rr
                            if (skipList[i]) {
                                t = getT4Len(q, -r);
                                anc2 = bezier(q, t);

                                if (!op && i == 0) {
                                    adjRdirAtEnd = t;
                                } else {
                                    pnts[pnts.length - 1][1] = adjHan(q[0], q[1], t);
                                }

                                ldir2 = defHan(t, q, 0);
                                rdir2 = getPnt(anc2, getRad(ldir2, anc2), r * hanLen);

                                pnts.push([anc2, rdir2, ldir2, ptyp]);
                            } else {
                                qb = [anc1, rdir1, adjHan(q[3], q[2], 1 - t), q[3]];
                                t = getT4Len(qb, -r);
                                anc2 = bezier(qb, t);
                                ldir2 = defHan(t, qb, 0);
                                rdir2 = getPnt(anc2, getRad(ldir2, anc2), r * hanLen);
                                rdir1 = adjHan(anc1, rdir1, t);

                                pnts.push([anc1, rdir1, ldir1, ptyp],
                                    [anc2, rdir2, ldir2, ptyp]);
                            }
                        }
                        redrawFlg = true;
                    }
                }
            }
            if (adjRdirAtEnd > 0) {
                pnts[pnts.length - 1][1] = adjHan(p[0].anchor, p[0].rightDirection, adjRdirAtEnd);
            }

            if (redrawFlg) {
                // redraw
                for (i = p.length - 1; i > 0; i--) p[i].remove();

                for (i = 0; i < pnts.length; i++) {
                    pt = i > 0 ? p.add() : p[0];
                    with (pt) {
                        anchor = pnts[i][0];
                        rightDirection = pnts[i][1];
                        leftDirection = pnts[i][2];
                        pointType = pnts[i][3];
                    }
                }
            }
        }
        activeDocument.selection = s;
        // alert(new Date() - tim);
    }

    // ------------------------------------------------
    // return [x,y] of the distance "len" and the angle "rad"(in radian)
    // from "pt"=[x,y]
    function getPnt(pt, rad, len) {
        return [pt[0] + Math.cos(rad) * len,
        pt[1] + Math.sin(rad) * len];
    }

    // ------------------------------------------------
    // return the [x, y] coordinate of the handle of the point on the bezier curve
    // that corresponds to the parameter "t"
    // n=0:leftDir, n=1:rightDir
    function defHan(t, q, n) {
        return [t * (t * (q[n][0] - 2 * q[n + 1][0] + q[n + 2][0]) + 2 * (q[n + 1][0] - q[n][0])) + q[n][0],
        t * (t * (q[n][1] - 2 * q[n + 1][1] + q[n + 2][1]) + 2 * (q[n + 1][1] - q[n][1])) + q[n][1]];
    }

    // -----------------------------------------------
    // return the [x, y] coordinate on the bezier curve
    // that corresponds to the paramter "t"
    function bezier(q, t) {
        var u = 1 - t;
        return [u * u * u * q[0][0] + 3 * u * t * (u * q[1][0] + t * q[2][0]) + t * t * t * q[3][0],
        u * u * u * q[0][1] + 3 * u * t * (u * q[1][1] + t * q[2][1]) + t * t * t * q[3][1]];
    }

    // ------------------------------------------------
    // adjust the length of the handle "dir"
    // by the magnification ratio "m",
    // returns the modified [x, y] coordinate of the handle
    // "anc" is the anchor [x, y]
    function adjHan(anc, dir, m) {
        return [anc[0] + (dir[0] - anc[0]) * m,
        anc[1] + (dir[1] - anc[1]) * m];
    }

    // ------------------------------------------------
    // return true if the pathPoints "p[idx]" is a corner
    function isCorner(p, idx) {
        var pnt0 = getAnglePnt(p, idx, -1);
        var pnt1 = getAnglePnt(p, idx, 1);
        if (!pnt0 || !pnt1) return false;                    // at the end of a open-path
        if (pnt0.length < 1 || pnt1.length < 1) return false;   // anchor is overlapping, so cannot determine the angle
        var rad = getRad2(pnt0, p[idx].anchor, pnt1, true);
        if (rad > Math.PI - 0.1) return false;   // set the angle tolerance here
        return true;
    }
    // ------------------------------------------------
    // "p"=pathPoints, "idx1"=index of pathpoint
    // "dir" = -1, returns previous point [x,y] to get the angle of tangent at pathpoints[idx1]
    // "dir" =  1, returns next ...
    function getAnglePnt(p, idx1, dir) {
        if (!dir) dir = -1;
        var idx2 = parseIdx(p, idx1 + dir);
        if (idx2 < 0) return null;  // at the end of a open-path
        var p2 = p[idx2];
        with (p[idx1]) {
            if (dir < 0) {
                if (arrEq(leftDirection, anchor)) {
                    if (arrEq(p2.anchor, anchor)) return [];
                    if (arrEq(p2.anchor, p2.rightDirection)
                        || arrEq(p2.rightDirection, anchor)) return p2.anchor;
                    else return p2.rightDirection;
                } else {
                    return leftDirection;
                }
            } else {
                if (arrEq(anchor, rightDirection)) {
                    if (arrEq(anchor, p2.anchor)) return [];
                    if (arrEq(p2.anchor, p2.leftDirection)
                        || arrEq(anchor, p2.leftDirection)) return p2.anchor;
                    else return p2.leftDirection;
                } else {
                    return rightDirection;
                }
            }
        }
    }
    // --------------------------------------
    // if the contents of both arrays are equal, return true (lengthes must be same)
    function arrEq(arr1, arr2) {
        for (var i = 0; i < arr1.length; i++) {
            if (arr1[i] != arr2[i]) return false;
        }
        return true;
    }

    // ------------------------------------------------
    // return the distance between p1=[x,y] and p2=[x,y]
    function dist(p1, p2) {
        return Math.sqrt(Math.pow(p1[0] - p2[0], 2)
            + Math.pow(p1[1] - p2[1], 2));
    }
    // ------------------------------------------------
    // return the squared distance between p1=[x,y] and p2=[x,y]
    function dist2(p1, p2) {
        return Math.pow(p1[0] - p2[0], 2)
            + Math.pow(p1[1] - p2[1], 2);
    }
    // --------------------------------------
    // return the angle in radian
    // of the line drawn from p1=[x,y] from p2
    function getRad(p1, p2) {
        return Math.atan2(p2[1] - p1[1],
            p2[0] - p1[0]);
    }

    // --------------------------------------
    // return the angle between two line segments
    // o-p1 and o-p2 ( 0 - Math.PI)
    function getRad2(p1, o, p2) {
        var v1 = normalize(p1, o);
        var v2 = normalize(p2, o);
        return Math.acos(v1[0] * v2[0] + v1[1] * v2[1]);
    }
    // ------------------------------------------------
    function normalize(p, o) {
        var d = dist(p, o);
        return d == 0 ? [0, 0] : [(p[0] - o[0]) / d,
        (p[1] - o[1]) / d];
    }

    // ------------------------------------------------
    // return the bezier curve parameter "t"
    // at the point which the length of the bezier curve segment
    // (from the point start drawing) is "len"
    // when "len" is 0, return the length of whole this segment.
    function getT4Len(q, len) {
        var m = [q[3][0] - q[0][0] + 3 * (q[1][0] - q[2][0]),
        q[0][0] - 2 * q[1][0] + q[2][0],
        q[1][0] - q[0][0]];
        var n = [q[3][1] - q[0][1] + 3 * (q[1][1] - q[2][1]),
        q[0][1] - 2 * q[1][1] + q[2][1],
        q[1][1] - q[0][1]];
        var k = [m[0] * m[0] + n[0] * n[0],
        4 * (m[0] * m[1] + n[0] * n[1]),
        2 * ((m[0] * m[2] + n[0] * n[2]) + 2 * (m[1] * m[1] + n[1] * n[1])),
        4 * (m[1] * m[2] + n[1] * n[2]),
        m[2] * m[2] + n[2] * n[2]];

        var fullLen = getLength(k, 1);

        if (len == 0) {
            return fullLen;

        } else if (len < 0) {
            len += fullLen;
            if (len < 0) return 0;

        } else if (len > fullLen) {
            return 1;
        }

        var t, d;
        var t0 = 0;
        var t1 = 1;
        var torelance = 0.001;

        for (var h = 1; h < 30; h++) {
            t = t0 + (t1 - t0) / 2;
            d = len - getLength(k, t);
            if (Math.abs(d) < torelance) break;
            else if (d < 0) t1 = t;
            else t0 = t;
        }
        return t;
    }

    // ------------------------------------------------
    // return the length of bezier curve segment
    // in range of parameter from 0 to "t"
    function getLength(k, t) {
        var h = t / 128;
        var hh = h * 2;
        var fc = function (t, k) {
            return Math.sqrt(t * (t * (t * (t * k[0] + k[1]) + k[2]) + k[3]) + k[4]) || 0
        };
        var total = (fc(0, k) - fc(t, k)) / 2;
        for (var i = h; i < t; i += hh) total += 2 * fc(i, k) + fc(i + h, k);
        return total * hh;
    }

    // --------------------------------------
    // merge nearly overlapped anchor points 
    // return the length of pathpoints after merging
    function readjustAnchors(p) {
        // Settings ==========================

        // merge the anchor points when the distance between
        // 2 points is within ### square root ### of this value (in point)
        var minDist = 0.0025;

        // ===================================
        if (p.length < 2) return 1;
        var i;

        if (p.parent.closed) {
            for (i = p.length - 1; i >= 1; i--) {
                if (dist2(p[0].anchor, p[i].anchor) < minDist) {
                    p[0].leftDirection = p[i].leftDirection;
                    p[i].remove();
                } else {
                    break;
                }
            }
        }

        for (i = p.length - 1; i >= 1; i--) {
            if (dist2(p[i].anchor, p[i - 1].anchor) < minDist) {
                p[i - 1].rightDirection = p[i].rightDirection;
                p[i].remove();
            }
        }

        return p.length;
    }
    // -----------------------------------------------
    // return pathpoint's index. when the argument is out of bounds,
    // fixes it if the path is closed (ex. next of last index is 0),
    // or return -1 if the path is not closed.
    function parseIdx(p, n) { // PathPoints, number for index
        var len = p.length;
        if (p.parent.closed) {
            return n >= 0 ? n % len : len - Math.abs(n % len);
        } else {
            return (n < 0 || n > len - 1) ? -1 : n;
        }
    }
    // -----------------------------------------------
    function getDat(p) { // pathPoint
        with (p) return [anchor, rightDirection, leftDirection, pointType];
    }
    // -----------------------------------------------
    function isSelected(p) { // PathPoint
        return p.selected == PathPointSelection.ANCHORPOINT;
    }

    function selectLeftCornersForRectangles(paths) {
        for (var i = 0; i < paths.length; i++) {
            var it = paths[i];
            if (!it || it.typename !== "PathItem" || !it.closed) continue;
            var pp = it.pathPoints;
            if (!pp || pp.length !== 4) continue;

            var sel = 0;
            for (var a = 0; a < 4; a++) if (pp[a].selected === PathPointSelection.ANCHORPOINT) sel++;
            if (sel > 0 && sel < 4) continue; // respect partial anchor selection

            var idx = [0, 1, 2, 3];
            idx.sort(function (x, y) { return pp[x].anchor[0] - pp[y].anchor[0]; });
            var a0 = idx[0], a1 = idx[1];
            var top = (pp[a0].anchor[1] >= pp[a1].anchor[1]) ? a0 : a1;
            var bot = (top === a0) ? a1 : a0;

            for (a = 0; a < 4; a++) pp[a].selected = PathPointSelection.NOSELECTION;
            pp[top].selected = PathPointSelection.ANCHORPOINT;
            pp[bot].selected = PathPointSelection.ANCHORPOINT;
        }
    }

})();