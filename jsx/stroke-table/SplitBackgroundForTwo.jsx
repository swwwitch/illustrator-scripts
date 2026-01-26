#target illustrator
#targetengine "SplitBackgroundForTwoEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：
SplitBackgroundForTwo

### 更新日：
20260126

### 概要：
2つのオブジェクト（テキスト、パス、グループなど）を選択して実行すると、各オブジェクトの背面に2分割の背景を作成します。
ダイアログで［方向］（左右／上下）を切り替えることで、左右2分割（Left/Right）だけでなく上下2分割（Top/Bottom）にも対応します。
実行時にサイズ倍率（%）を指定するダイアログが表示され、閉じる前にプレビューを確認できます（デフォルトは200%）。

分割の基準となるギャップは、2つのオブジェクト間の「隙間」を基準に計算されます。
［バランス］パネルで「なし／左／右（上下モード時は なし／上／下）」を選択すると、
- 「なし」：均等（中央分割）
- 「左（上）」：左（上）側のマージンを［幅］で指定（反対側は自動計算）
- 「右（下）」：右（下）側のマージンを［幅］で指定（反対側は自動計算）

［幅］はスライダーおよび数値入力で指定でき、選択中の2オブジェクト間ギャップを最大値として自動設定されます。
テキストのサイドベアリング等によるズレを減らすため、
一時的にテキストをアウトライン化して外接矩形を計算し、計算後すぐに一時生成物を削除します。
その上で背景長方形を選択オブジェクトの背面に配置し、ダイアログ内でリアルタイムにプレビュー表示します（元のオブジェクトは変更しません）。

サイズ（%）および［幅］入力欄では、↑↓キーで±1、Shift+↑↓で±10（10刻みスナップ）、Option+↑↓で±0.1 の増減が可能です。

### 更新履歴：
- v1.0 (20260124) : 初期バージョン
- v1.1 (20260126) : ［バランス］（なし／左／右）と［幅］指定による左右背景の比率調整に対応。［幅］はオブジェクト間ギャップを最大値として自動計算し、スライダー／数値入力／矢印キー操作に対応
- v2.0 (20260126) : ［方向］（左右／上下）に対応。上下配置時は背景を上下2分割で作成し、バランス（なし／上／下）と幅でギャップ内の分割位置を調整可能。
- v2.1 (20260126) : 選択オブジェクトの位置関係から左右／上下を自動判別し、ダイアログの［方向］UIを省略。
- v2.2 (20260126) : #targetengine を追加し、Illustrator再起動までは前回閉じたダイアログ値を復元。
- v2.3 (20260126) : PreviewHistory によるプレビュー一括Undoを追加（プレビューでヒストリーを汚さない）。
- v2.4 (20260126) : ［OK］時にプレビューUndo後も必ず最終描画を行い、直前の見た目で確定するよう修正。
- v2.5 (20260126) : ［OK］時にプレビュー削除で last 値が初期化される不具合を修正（角丸等が反映されない問題）。
- v2.6 (20260126) : PreviewHistory.undo() 後に一時アウトライン生成物が復活するケースを除去（TEMPマーカーを即掃除）。
*/

// --- Version / バージョン ---
var SCRIPT_VERSION = "v2.6";

// --- Dialog UI prefs / ダイアログUI設定 ---
var DIALOG_OFFSET_X = 300;
var DIALOG_OFFSET_Y = 0;
var DIALOG_OPACITY = 0.98;

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

    // Direction
    panelDirection: { ja: "方向", en: "Direction" },
    directionHorizontal: { ja: "左右", en: "Left/Right" },
    directionVertical: { ja: "上下", en: "Top/Bottom" },

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
    // vertical-mode main size label (re-uses the same percent field)
    labelWidthMain: {
        ja: "幅：",
        en: "Width:"
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
    // vertical-mode fill labels
    labelFillTop: {
        ja: "塗り（上）",
        en: "Fill (Top)"
    },
    labelFillBottom: {
        ja: "塗り（下）",
        en: "Fill (Bottom)"
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
    // vertical-mode balance labels
    fixedTop: {
        ja: "上",
        en: "Top"
    },
    fixedBottom: {
        ja: "下",
        en: "Bottom"
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
        ja: "サイズ（%）は0より大きい数値を入力してください。",
        en: "Enter a size (%) greater than 0."
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

// --- Session settings (kept until Illustrator restart) / セッション設定（Ai再起動まで保持） ---
var SETTINGS_KEY = "SplitBackgroundForTwo_settings";

function getSessionSettings() {
    try {
        if (!$.global[SETTINGS_KEY]) {
            $.global[SETTINGS_KEY] = {
                percent: 200,
                preview: true,
                overallFrame: false,
                divider: false,
                fillLeft: true,
                fillRight: true,
                strokeUnit: formatUnitValue(ptToUnit(1, "strokeUnits")),
                cornerUnit: formatUnitValue(ptToUnit(0, "rulerType")),
                balanceMode: 'none',
                widthUnit: 0
            };
        }
        return $.global[SETTINGS_KEY];
    } catch (e) {
        return {
            percent: 200,
            preview: true,
            overallFrame: false,
            divider: false,
            fillLeft: true,
            fillRight: true,
            strokeUnit: formatUnitValue(ptToUnit(1, "strokeUnits")),
            cornerUnit: formatUnitValue(ptToUnit(0, "rulerType")),
            balanceMode: 'none',
            widthUnit: 0
        };
    }
}

function saveSessionSettings(s) {
    try { $.global[SETTINGS_KEY] = s; } catch (e) { }
}

(function () {
    if (app.documents.length === 0) {
        alert(L("alertOpenDoc"));
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (!sel || sel.length !== 2) {
        alert(L("alertSelectTwoItems"));
        return;
    }

    // Keep original two items; ordering depends on direction at draw time
    var itemA = sel[0];
    var itemB = sel[1];

    // --- Auto direction detection / 方向を自動判別 ---
    function detectDirectionModeBySelection(itemA, itemB) {
        var mode = 'horizontal';
        var tempGroup = null;
        try {
            var hostLayer = null;
            try { hostLayer = itemA.layer; } catch (e0) { hostLayer = null; }
            if (!hostLayer) {
                try { hostLayer = doc.activeLayer; } catch (e1) { hostLayer = null; }
            }
            if (!hostLayer) hostLayer = doc.layers[0];

            unlockAndShow(hostLayer);
            tempGroup = hostLayer.groupItems.add();
            tempGroup.name = TEMP_NAME_PREFIX + "__temp_outline_group_for_direction__";
            markTemp(tempGroup);
            try { tempGroup.opacity = 0; } catch (eOp) { }

            var bi1 = getBoundsInfo(itemA, tempGroup);
            var bi2 = getBoundsInfo(itemB, tempGroup);
            var b1 = bi1.bounds;
            var b2 = bi2.bounds;

            // Horizontal gap
            var leftB = (b1[0] <= b2[0]) ? b1 : b2;
            var rightB = (b1[0] <= b2[0]) ? b2 : b1;
            var gapH = rightB[0] - leftB[2];
            if (gapH < 0) gapH = 0;

            // Vertical gap
            var upperB = (b1[1] >= b2[1]) ? b1 : b2;
            var lowerB = (b1[1] >= b2[1]) ? b2 : b1;
            var gapV = upperB[3] - lowerB[1];
            if (gapV < 0) gapV = 0;

            // Decide: larger gap wins (tie => horizontal)
            mode = (gapV > gapH) ? 'vertical' : 'horizontal';
        } catch (e) {
            mode = 'horizontal';
        } finally {
            try { removeMarkedTempItems(doc); } catch (e2) { }
            try {
                if (tempGroup && tempGroup.isValid) {
                    unlockAndShow(tempGroup);
                    try { tempGroup.remove(); } catch (e3) { }
                }
            } catch (e4) { }
            try { removeMarkedTempItems(doc); } catch (e5) { }
        }
        return mode;
    }

    // Decide once per run
    var AUTO_DIRECTION_MODE = detectDirectionModeBySelection(itemA, itemB);
    lastDirectionMode = AUTO_DIRECTION_MODE;

    // --- Temporary marker (cleanup) / 一時生成物マーカー（後片付け） ---
    var TEMP_NOTE = "__SplitBackgroundForTwo_TEMP__";
    var TEMP_NAME_PREFIX = "__SplitBackgroundForTwo_TEMP__";

    // --- Preview marker (cleanup) / プレビューマーカー（後片付け） ---
    var PREVIEW_NOTE = "__SplitBackgroundForTwo_PREVIEW__";

    /* =========================================
 * PreviewHistory util (extractable)
 * ヒストリーを残さないプレビューのための小さなユーティリティ。
 * 他スクリプトでもこのブロックをコピペすれば再利用できます。
 * 使い方:
 *   PreviewHistory.start();      // ダイアログ表示時などにカウンタ初期化
 *   PreviewHistory.bump();       // プレビュー描画ごとにカウント(+1)
 *   PreviewHistory.undo();       // 閉じる/キャンセル時に一括Undo
 *   PreviewHistory.cancelTask(t);// app.scheduleTaskのキャンセル補助
 * ========================================= */
    (function (g) {
        if (!g.PreviewHistory) {
            g.PreviewHistory = {
                start: function () { g.__previewUndoCount = 0; },
                bump: function () { g.__previewUndoCount = (g.__previewUndoCount | 0) + 1; },
                undo: function () {
                    var n = g.__previewUndoCount | 0;
                    try { for (var i = 0; i < n; i++) app.executeMenuCommand('undo'); } catch (e) { }
                    g.__previewUndoCount = 0;
                },
                cancelTask: function (taskId) {
                    try { if (taskId) app.cancelTask(taskId); } catch (e) { }
                }
            };
        }
    })($.global);

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

    function unlockAndShow(item) {
        if (!item) return;
        try { item.locked = false; } catch (e1) { }
        try { item.hidden = false; } catch (e2) { }
        try { item.visible = true; } catch (e3) { }
        try { item.template = false; } catch (e4) { }
        try { item.printable = true; } catch (e5) { }
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

    function markTemp(item) {
        if (!item) return;

        // note（従来）
        try { item.note = TEMP_NOTE; } catch (e0) { }

        // name（Undoでnoteが失われる対策）
        try {
            if (item.name !== undefined) {
                var nm = "";
                try { nm = String(item.name || ""); } catch (eN0) { nm = ""; }
                if (nm.indexOf(TEMP_NAME_PREFIX) !== 0) {
                    item.name = TEMP_NAME_PREFIX;
                }
            }
        } catch (eName) { }

        // 子要素もマーキング
        try {
            if (item.pageItems && item.pageItems.length) {
                for (var i = item.pageItems.length - 1; i >= 0; i--) {
                    markTemp(item.pageItems[i]);
                }
            }
        } catch (eChildren) { }
    }

    function removeMarkedTempItems(doc) {
        if (!doc) return;

        // pageItems: note または name プレフィックスで消す
        try {
            for (var i = doc.pageItems.length - 1; i >= 0; i--) {
                var it = doc.pageItems[i];
                var isMarked = false;

                try { isMarked = (it.note === TEMP_NOTE); } catch (eNote) { isMarked = false; }

                if (!isMarked) {
                    try {
                        if (it.name !== undefined) {
                            var nm = String(it.name || "");
                            isMarked = (nm.indexOf(TEMP_NAME_PREFIX) === 0);
                        }
                    } catch (eNm) { }
                }

                if (isMarked) {
                    unlockAndShow(it);
                    try { it.remove(); } catch (eRm) { }
                }
            }
        } catch (eLoop) { }

        // groupItems: 念のため group 名でも掃除（Undo後に漏れる保険）
        try {
            for (var g = doc.groupItems.length - 1; g >= 0; g--) {
                var gi = doc.groupItems[g];
                var hit = false;
                try {
                    if (gi.name !== undefined) {
                        var gnm = String(gi.name || "");
                        hit = (gnm.indexOf(TEMP_NAME_PREFIX) === 0) || (gnm.indexOf("__temp_outline_group") === 0);
                    }
                } catch (eGnm) { }
                if (hit) {
                    try { unlockAndShow(gi); } catch (eUG) { }
                    try { gi.remove(); } catch (eRG) { }
                }
            }
        } catch (eGrp) { }
    }

    /**
     * テキストを「一時グループ」へ複製→アウトライン化し、その外接矩形を返す。
     */
    function getOutlineBoundsFromText(tf, tempGroup) {
        var dup = null;
        var outlineItem = null;
        var b;

        try {
            dup = tf.duplicate(tempGroup, ElementPlacement.PLACEATBEGINNING);
            markTemp(dup);
            unlockAndShow(dup);

            try {
                outlineItem = dup.createOutline();
                if (outlineItem) {
                    markTemp(outlineItem);
                    unlockAndShow(outlineItem);
                    try {
                        if (outlineItem.parent !== tempGroup) {
                            outlineItem.move(tempGroup, ElementPlacement.PLACEATBEGINNING);
                        }
                    } catch (eMove) { }
                }
            } catch (eOutline) {
                outlineItem = null;
            }

            var target = outlineItem ? outlineItem : dup;
            b = target.geometricBounds; // [left, top, right, bottom]

        } finally {
            // ★ 取得後に即削除（残留防止）
            try {
                if (outlineItem && outlineItem.isValid) {
                    unlockAndShow(outlineItem);
                    outlineItem.remove();
                }
            } catch (eRmO) { }
            try {
                if (dup && dup.isValid) {
                    unlockAndShow(dup);
                    dup.remove();
                }
            } catch (eRmD) { }
        }

        if (!b || b.length < 4) b = [0, 0, 0, 0];

        return {
            bounds: [b[0], b[1], b[2], b[3]],
            outline: null,
            duplicateText: null
        };
    }

    /**
     * 任意オブジェクトの外接矩形（geometricBounds）を取得する。
     * TextFrame の場合はアウトライン化した複製から取得して精度を上げる。
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

        var dup = null;
        var b;
        try {
            dup = item.duplicate(tempGroup, ElementPlacement.PLACEATBEGINNING);
            markTemp(dup);
            unlockAndShow(dup);
            b = dup.geometricBounds;
        } finally {
            try {
                if (dup && dup.isValid) {
                    unlockAndShow(dup);
                    dup.remove();
                }
            } catch (eRmDup) { }
        }

        if (!b || b.length < 4) b = [0, 0, 0, 0];
        return {
            bounds: [b[0], b[1], b[2], b[3]],
            outline: null,
            duplicateItem: null
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
    colorLeft.red = 220; colorLeft.green = 220; colorLeft.blue = 220;

    var colorRight = new RGBColor();
    colorRight.red = 128; colorRight.green = 128; colorRight.blue = 128;

    // K100 (CMYK) for frame / 枠線用 K100
    var colorK100 = new CMYKColor();
    colorK100.cyan = 0;
    colorK100.magenta = 0;
    colorK100.yellow = 0;
    colorK100.black = 100;

    // Round Corners Live Effect (for Overall Frame only)
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

    // --- Fill corner rounding / 塗りの角丸（選択アンカーのみ丸める） ---
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
            if (sel > 0 && sel < 4) continue;

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

    function selectTopCornersForRectangles(paths) {
        for (var i = 0; i < paths.length; i++) {
            var it = paths[i];
            if (!it || it.typename !== "PathItem" || !it.closed) continue;
            var pp = it.pathPoints;
            if (!pp || pp.length !== 4) continue;

            var sel = 0;
            for (var a = 0; a < 4; a++) if (pp[a].selected === PathPointSelection.ANCHORPOINT) sel++;
            if (sel > 0 && sel < 4) continue;

            var idx = [0, 1, 2, 3];
            idx.sort(function (x, y) { return pp[y].anchor[1] - pp[x].anchor[1]; }); // top first
            var a0 = idx[0], a1 = idx[1];

            for (a = 0; a < 4; a++) pp[a].selected = PathPointSelection.NOSELECTION;
            pp[a0].selected = PathPointSelection.ANCHORPOINT;
            pp[a1].selected = PathPointSelection.ANCHORPOINT;
        }
    }

    function selectBottomCornersForRectangles(paths) {
        for (var i = 0; i < paths.length; i++) {
            var it = paths[i];
            if (!it || it.typename !== "PathItem" || !it.closed) continue;
            var pp = it.pathPoints;
            if (!pp || pp.length !== 4) continue;

            var sel = 0;
            for (var a = 0; a < 4; a++) if (pp[a].selected === PathPointSelection.ANCHORPOINT) sel++;
            if (sel > 0 && sel < 4) continue;

            var idx = [0, 1, 2, 3];
            idx.sort(function (x, y) { return pp[x].anchor[1] - pp[y].anchor[1]; }); // bottom first
            var a0 = idx[0], a1 = idx[1];

            for (a = 0; a < 4; a++) pp[a].selected = PathPointSelection.NOSELECTION;
            pp[a0].selected = PathPointSelection.ANCHORPOINT;
            pp[a1].selected = PathPointSelection.ANCHORPOINT;
        }
    }

    function applyFillCornerRounding(pathItem, side, radiusPt) {
        if (!pathItem) return;
        var r = Number(radiusPt);
        if (isNaN(r) || r <= 0) return;

        clearAnchorSelection(pathItem);

        if (side === "left") {
            selectLeftCornersForRectangles([pathItem]);
        } else if (side === "right") {
            selectRightCornersForRectangles([pathItem]);
        } else if (side === "top") {
            selectTopCornersForRectangles([pathItem]);
        } else if (side === "bottom") {
            selectBottomCornersForRectangles([pathItem]);
        } else {
            return;
        }

        try {
            roundAnyCorner([pathItem], { rr: r });
        } catch (e) { }

        clearAnchorSelection(pathItem);
    }

    // --- 背景を置くターゲットレイヤーの決定 ---
    function getLayerIndex(doc, layer) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i] === layer) return i;
        }
        return -1;
    }

    // Illustratorのlayers配列は基本的に「上（前面）→下（背面）」順で並ぶ想定。
    function getBackmostLayer(doc, layerA, layerB) {
        var ia = getLayerIndex(doc, layerA);
        var ib = getLayerIndex(doc, layerB);
        if (ia < 0) return layerB;
        if (ib < 0) return layerA;
        return (ia > ib) ? layerA : layerB;
    }

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
    var lastDirectionMode = 'horizontal';

    function clearPreviewRects() {
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
     * アウトライン計算→背景長方形を作成
     * directionMode: 'horizontal' | 'vertical'
     */
    function drawBackgroundRects(itemA0, itemB0, sizeRatioValue, addOverallFrame, addDivider, addFillLeft, addFillRight, strokeWidthValue, cornerRadius, balanceMode, widthPercent, directionMode) {
        directionMode = directionMode || 'horizontal';

        // --- 一時グループ（アウトライン生成用）---
        var originalActiveLayer = doc.activeLayer;

        var hostLayer = null;
        try { hostLayer = itemA0.layer; } catch (eHL) { hostLayer = null; }
        if (!hostLayer) {
            try { hostLayer = doc.activeLayer; } catch (eHL2) { hostLayer = null; }
        }
        if (!hostLayer) hostLayer = doc.layers[0];

        unlockAndShow(hostLayer);
        var tempGroup = hostLayer.groupItems.add();

        markTemp(tempGroup);
        try { tempGroup.opacity = 0; } catch (eOp) { }
        tempGroup.name = TEMP_NAME_PREFIX + "__temp_outline_group__";
        // 計算結果
        var rectTop, rectHeight;
        var rectLeftStart, gapCenter, rectRightEnd;

        // items ordering will be decided after outline bounds are obtained
        var firstItem = itemA0;
        var secondItem = itemB0;

        var b1, b2;

        try {
            var outline1Info = getBoundsInfo(itemA0, tempGroup);
            var outline2Info = getBoundsInfo(itemB0, tempGroup);

            b1 = outline1Info.bounds;
            b2 = outline2Info.bounds;

            // Decide ordering by outline bounds (more accurate for TextFrame)
            if (directionMode === 'vertical') {
                // first = upper (higher top)
                if (b1[1] < b2[1]) {
                    firstItem = itemB0; secondItem = itemA0;
                    var tb = b1; b1 = b2; b2 = tb;
                }
            } else {
                // first = left (smaller left)
                if (b1[0] > b2[0]) {
                    firstItem = itemB0; secondItem = itemA0;
                    var xb = b1; b1 = b2; b2 = xb;
                }
            }

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

            function clamp(v, mn, mx) {
                if (v < mn) return mn;
                if (v > mx) return mx;
                return v;
            }

            var wPt = Number(widthPercent);
            if (isNaN(wPt) || wPt < 0) wPt = 0;

            if (directionMode === 'vertical') {
                // ===== Vertical arrangement: split top/bottom =====
                var top1 = b1[1], bot1 = b1[3], left1 = b1[0], right1 = b1[2];
                var top2 = b2[1], bot2 = b2[3], left2 = b2[0], right2 = b2[2];

                // Horizontal size (scaled by sizeRatioValue)
                var leftLimit = Math.min(left1, left2);
                var rightLimit = Math.max(right1, right2);
                var contentW = rightLimit - leftLimit;
                if (contentW < 0) contentW = 0;

                var centerX = leftLimit + (contentW / 2);
                var rectW = contentW * sizeRatioValue;
                var rectL = centerX - (rectW / 2);
                var rectR = rectL + rectW;

                // Vertical split based on gap
                var gapDist = bot1 - top2;
                if (gapDist < 0) gapDist = 0;

                var marginTop = gapDist / 2;
                var marginBottom = gapDist / 2;

                if (balanceMode === 'left') {
                    // 'left' UI becomes 'top'
                    marginTop = clamp(wPt, 0, gapDist);
                    marginBottom = gapDist - marginTop;
                } else if (balanceMode === 'right') {
                    // 'right' UI becomes 'bottom'
                    marginBottom = clamp(wPt, 0, gapDist);
                    marginTop = gapDist - marginBottom;
                } else {
                    marginTop = gapDist / 2;
                    marginBottom = gapDist / 2;
                }

                rectTop = top1 + marginTop;
                var rectBottom = bot2 - marginBottom;
                rectHeight = rectTop - rectBottom;
                if (rectHeight < 0) rectHeight = 0;

                // Split boundary Y inside the gap
                var splitY = bot1 - marginTop;

                rectLeftStart = rectL;
                rectRightEnd = rectR;
                gapCenter = splitY; // reuse as split position (Y)

            } else {
                // ===== Horizontal arrangement: split left/right (existing behavior) =====
                var topLimit = Math.max(b1[1], b2[1]);
                var bottomLimit = Math.min(b1[3], b2[3]);

                var contentHeight = topLimit - bottomLimit;
                var centerY = topLimit - (contentHeight / 2);

                rectHeight = contentHeight * sizeRatioValue;
                rectTop = centerY + (rectHeight / 2);

                var item1Left = b1[0];
                var item1Right = b1[2];
                var item2Left = b2[0];
                var item2Right = b2[2];

                var gapDistH = item2Left - item1Right;
                if (gapDistH < 0) gapDistH = 0;

                var marginLeft = gapDistH / 2;
                var marginRight = gapDistH / 2;

                if (balanceMode === 'left') {
                    marginLeft = clamp(wPt, 0, gapDistH);
                    marginRight = gapDistH - marginLeft;
                } else if (balanceMode === 'right') {
                    marginRight = clamp(wPt, 0, gapDistH);
                    marginLeft = gapDistH - marginRight;
                } else {
                    marginLeft = gapDistH / 2;
                    marginRight = gapDistH / 2;
                }

                gapCenter = item1Right + marginLeft;
                rectLeftStart = item1Left - marginLeft;
                rectRightEnd = item2Right + marginRight;
            }

        } finally {
            try { doc.selection = null; } catch (eSel) { }

            try {
                if (originalActiveLayer && originalActiveLayer.isValid) {
                    doc.activeLayer = originalActiveLayer;
                }
            } catch (eRestore) { }

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

            removeMarkedTempItems(doc);
        }

        // --- ここから描画（アウトラインは既に削除済み）---
        var leftLayer = null;
        var rightLayer = null;
        try { leftLayer = firstItem.layer; } catch (eLL) { leftLayer = null; }
        try { rightLayer = secondItem.layer; } catch (eRL) { rightLayer = null; }

        var rect1 = null;
        var rect2 = null;

        if (directionMode === 'vertical') {
            // Top / Bottom rectangles
            var rectBottomAll = rectTop - rectHeight;

            if (addFillLeft) {
                // Top
                rect1 = (leftLayer ? leftLayer.pathItems : doc.pathItems).rectangle(
                    rectTop,
                    rectLeftStart,
                    rectRightEnd - rectLeftStart,
                    rectTop - gapCenter
                );
                rect1.fillColor = colorLeft;
                rect1.stroked = false;
            }

            if (addFillRight) {
                // Bottom
                rect2 = (rightLayer ? rightLayer.pathItems : doc.pathItems).rectangle(
                    gapCenter,
                    rectLeftStart,
                    rectRightEnd - rectLeftStart,
                    gapCenter - rectBottomAll
                );
                rect2.fillColor = colorRight;
                rect2.stroked = false;
            }
        } else {
            // Left / Right rectangles
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
        }

        // Fill corner rounding
        if (cornerRadius && Number(cornerRadius) > 0) {
            if (directionMode === 'vertical') {
                if (rect1) applyFillCornerRounding(rect1, "top", cornerRadius);
                if (rect2) applyFillCornerRounding(rect2, "bottom", cornerRadius);
            } else {
                if (rect1) applyFillCornerRounding(rect1, "left", cornerRadius);
                if (rect2) applyFillCornerRounding(rect2, "right", cornerRadius);
            }
        }

        var frameRect = null;
        if (addOverallFrame) {
            try {
                var frameLayer = getBackmostLayer(doc, leftLayer, rightLayer) || leftLayer || rightLayer;
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

                applyRoundCornersIfNeeded(frameRect, cornerRadius);

                try { frameRect.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (eFZ) { }
            } catch (eFrame) {
                frameRect = null;
            }
        }

        var dividerLine = null;
        if (addDivider) {
            try {
                var dividerLayer = getBackmostLayer(doc, leftLayer, rightLayer) || leftLayer || rightLayer;
                dividerLine = (dividerLayer ? dividerLayer.pathItems : doc.pathItems).add();
                dividerLine.filled = false;
                dividerLine.stroked = true;
                dividerLine.strokeWidth = strokeWidthValue;
                dividerLine.strokeColor = colorK100;

                if (directionMode === 'vertical') {
                    dividerLine.setEntirePath([
                        [rectLeftStart, gapCenter],
                        [rectRightEnd, gapCenter]
                    ]);
                } else {
                    dividerLine.setEntirePath([
                        [gapCenter, rectTop],
                        [gapCenter, rectTop - rectHeight]
                    ]);
                }

                try { dividerLine.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (eDZ) { }
            } catch (eDiv) {
                dividerLine = null;
            }
        }

        try { if (rect1) rect1.zOrder(ZOrderMethod.SENDTOBACK); } catch (eZ1) { }
        try { if (rect2) rect2.zOrder(ZOrderMethod.SENDTOBACK); } catch (eZ2) { }

        try { firstItem.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (eT1) { }
        try { secondItem.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (eT2) { }

        safeRedraw();
        return { rect1: rect1, rect2: rect2, frame: frameRect, divider: dividerLine };
    }

    /**
     * ↑↓キーで数値を増減するユーティリティ
     */
    function changeValueByArrowKey(editText, allowNegative, onChange) {
        if (!editText) return;

        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                } else if (event.keyName === "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                }
                event.preventDefault();
            } else if (keyboard.altKey) {
                delta = 0.1;
                if (event.keyName === "Up") value += delta;
                else if (event.keyName === "Down") value -= delta;
                event.preventDefault();
            } else {
                delta = 1;
                if (event.keyName === "Up") value += delta;
                else if (event.keyName === "Down") value -= delta;
                event.preventDefault();
            }

            if (keyboard.altKey) value = Math.round(value * 10) / 10;
            else value = Math.round(value);

            if (!allowNegative && value < 0) value = 0;

            editText.text = String(value);

            if (typeof onChange === "function") {
                try { onChange(); } catch (e) { }
            }
        });
    }

    /**
     * ダイアログ（プレビュー付き）
     */
    function showHeightDialog(defaultPercent, previewFn, clearPreviewFn) {
        var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
        dlg.orientation = 'column';
        dlg.alignChildren = ['fill', 'top'];
        dlg.margins = 18;

        try { PreviewHistory.start(); } catch (ePH0) { }

        var ss = getSessionSettings();

        setDialogOpacity(dlg, DIALOG_OPACITY);
        shiftDialogPosition(dlg, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);

        var row = dlg.add('group');
        row.orientation = 'row';
        row.alignChildren = ['left', 'center'];
        row.alignment = ['center', 'top'];

        // Main size label (will be swapped by direction)
        var stMainSizeLabel = row.add('statictext', undefined, L('labelHeight'));

        var et = row.add('edittext', undefined, String((ss && ss.percent !== undefined) ? ss.percent : defaultPercent));
        et.characters = 6;
        et.active = true;
        changeValueByArrowKey(et, false, applyPreview);

        row.add('statictext', undefined, L('labelPercent'));

        // --- Direction UI removed; mode is fixed to AUTO_DIRECTION_MODE ---
        function getDirectionMode() {
            return AUTO_DIRECTION_MODE;
        }

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

        // Keep variable names; label text will be swapped for vertical mode
        var cbFillLeft = drawPanel.add('checkbox', undefined, L('labelFillLeft'));
        cbFillLeft.value = (ss.fillLeft !== undefined) ? !!ss.fillLeft : true;

        var cbFillRight = drawPanel.add('checkbox', undefined, L('labelFillRight'));
        cbFillRight.value = (ss.fillRight !== undefined) ? !!ss.fillRight : true;

        var cbOverallFrame = drawPanel.add('checkbox', undefined, L('labelOverallFrame'));
        cbOverallFrame.value = (ss.overallFrame !== undefined) ? !!ss.overallFrame : false;

        var cbDivider = drawPanel.add('checkbox', undefined, L('labelDivider'));
        cbDivider.value = (ss.divider !== undefined) ? !!ss.divider : false;

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
        var etStroke = lineRow.add('edittext', undefined, (ss && ss.strokeUnit !== undefined) ? String(ss.strokeUnit) : formatUnitValue(ptToUnit(1, "strokeUnits")));
        etStroke.characters = 4;
        changeValueByArrowKey(etStroke, false, applyPreview);
        lineRow.add('statictext', undefined, getCurrentUnitLabelByPrefKey("strokeUnits"));

        var cornerRow = linePanel.add('group');
        cornerRow.orientation = 'row';
        cornerRow.alignChildren = ['left', 'center'];

        cornerRow.add('statictext', undefined, L('labelCornerRadius'));
        var etCorner = cornerRow.add('edittext', undefined, (ss && ss.cornerUnit !== undefined) ? String(ss.cornerUnit) : formatUnitValue(ptToUnit(0, "rulerType")));
        etCorner.characters = 4;
        cornerRow.add('statictext', undefined, getCurrentUnitLabelByPrefKey("rulerType"));
        changeValueByArrowKey(etCorner, false, applyPreview);

        function updateStrokeWidthEnabled() {
            var on = !!(cbOverallFrame.value || cbDivider.value);
            try { lineRow.enabled = on; } catch (eEn1) { }
        }

        function updateCornerEnabled() {
            var on = !!((cbFillLeft.value || cbFillRight.value) || cbOverallFrame.value);
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

        // Keep variable names; text will be swapped for vertical mode
        var rbPinNone = pinRadioRow.add('radiobutton', undefined, L('fixedNone'));
        var rbPinLeft = pinRadioRow.add('radiobutton', undefined, L('fixedLeft'));
        var rbPinRight = pinRadioRow.add('radiobutton', undefined, L('fixedRight'));
        rbPinNone.value = (ss.balanceMode === 'none');
        rbPinLeft.value = (ss.balanceMode === 'left');
        rbPinRight.value = (ss.balanceMode === 'right');
        if (!rbPinNone.value && !rbPinLeft.value && !rbPinRight.value) rbPinNone.value = true;

        // --- 幅 / Width (inside Balance panel) ---
        var pinWidthCol = pinPanel.add('group');
        pinWidthCol.orientation = 'column';
        pinWidthCol.alignChildren = ['fill', 'top'];
        pinWidthCol.spacing = 6;

        var pinWidthValueRow = pinWidthCol.add('group');
        pinWidthValueRow.orientation = 'row';
        pinWidthValueRow.alignChildren = ['left', 'center'];

        pinWidthValueRow.add('statictext', undefined, L('labelWidth'));

        var defaultWidthUnit = (ss && ss.widthUnit !== undefined)
            ? ss.widthUnit
            : ((AUTO_DIRECTION_MODE === 'vertical') ? 130 : 0);
        var etWidth = pinWidthValueRow.add('edittext', undefined, String(defaultWidthUnit));
        etWidth.characters = 6;
        changeValueByArrowKey(etWidth, false, applyPreview);

        pinWidthValueRow.add('statictext', undefined, getCurrentUnitLabelByPrefKey("rulerType"));

        var pinWidthSliderRow = pinWidthCol.add('group');
        pinWidthSliderRow.orientation = 'row';
        pinWidthSliderRow.alignChildren = ['fill', 'center'];

        var slWidth = pinWidthSliderRow.add('slider', undefined, 0, 0, 0);
        slWidth.preferredSize = [220, 20];

        function clampWidthValue(v, forceInteger) {
            v = Number(v);
            if (isNaN(v)) return 0;
            if (v < 0) v = 0;

            var maxV = 0;
            try { maxV = Number(slWidth.maxvalue); } catch (eMax0) { maxV = 0; }
            if (isNaN(maxV) || maxV < 0) maxV = 0;

            if (v > maxV) v = maxV;

            if (forceInteger) return Math.round(v);
            return Math.round(v * 10) / 10;
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

        function updateWidthSliderMaxBySelection(directionMode) {
            var maxUnit = 0;
            var tempGroupForMax = null;

            try {
                var hostLayer2 = null;
                try { hostLayer2 = itemA.layer; } catch (eHLm) { hostLayer2 = null; }
                if (!hostLayer2) {
                    try { hostLayer2 = doc.activeLayer; } catch (eHLm2) { hostLayer2 = null; }
                }
                if (!hostLayer2) hostLayer2 = doc.layers[0];

                unlockAndShow(hostLayer2);
                tempGroupForMax = hostLayer2.groupItems.add();
                tempGroupForMax.name = TEMP_NAME_PREFIX + "__temp_outline_group_for_max__";
                markTemp(tempGroupForMax);
                try { tempGroupForMax.opacity = 0; } catch (eOpM) { }

                var bi1 = getBoundsInfo(itemA, tempGroupForMax);
                var bi2 = getBoundsInfo(itemB, tempGroupForMax);
                var b1 = bi1.bounds;
                var b2 = bi2.bounds;

                var gapPt = 0;
                var dirMode = directionMode || AUTO_DIRECTION_MODE;

                if (dirMode === 'vertical') {
                    // upper = higher top
                    var upperB = (b1[1] >= b2[1]) ? b1 : b2;
                    var lowerB = (b1[1] >= b2[1]) ? b2 : b1;

                    var upperBottom = upperB[3];
                    var lowerTop = lowerB[1];
                    gapPt = upperBottom - lowerTop;
                } else {
                    var leftB = (b1[0] <= b2[0]) ? b1 : b2;
                    var rightB = (b1[0] <= b2[0]) ? b2 : b1;

                    var leftRight = leftB[2];
                    var rightLeft = rightB[0];
                    gapPt = rightLeft - leftRight;
                }

                if (gapPt < 0) gapPt = 0;

                maxUnit = ptToUnit(gapPt, "rulerType");
                if (isNaN(maxUnit) || maxUnit < 0) maxUnit = 0;
                maxUnit = Math.round(maxUnit * 10) / 10;

            } catch (eMax) {
                maxUnit = 0;
            } finally {
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
            try { if (slWidth.maxvalue < 0) slWidth.maxvalue = 0; } catch (eSetMax2) { }

            syncWidthUIFromEdit();
        }

        slWidth.onChanging = function () {
            syncWidthUIFromSlider();
            applyPreview();
        };

        etWidth.addEventListener('changing', function () {
            syncWidthUIFromEdit();
            applyPreview();
        });

        function updateWidthEnabled() {
            try { pinWidthCol.enabled = !rbPinNone.value; } catch (e) { }
        }

        rbPinNone.onClick = function () { updateWidthEnabled(); applyPreview(); };
        rbPinLeft.onClick = function () { updateWidthEnabled(); applyPreview(); };
        rbPinRight.onClick = function () { updateWidthEnabled(); applyPreview(); };

        updateWidthEnabled();

        function updateDirectionUILabels() {
            var isV = (AUTO_DIRECTION_MODE === 'vertical');

            try { stMainSizeLabel.text = isV ? L('labelWidthMain') : L('labelHeight'); } catch (e0) { }

            try { cbFillLeft.text = isV ? L('labelFillTop') : L('labelFillLeft'); } catch (e1) { }
            try { cbFillRight.text = isV ? L('labelFillBottom') : L('labelFillRight'); } catch (e2) { }

            try { rbPinLeft.text = isV ? L('fixedTop') : L('fixedLeft'); } catch (e3) { }
            try { rbPinRight.text = isV ? L('fixedBottom') : L('fixedRight'); } catch (e4) { }
        }

        updateDirectionUILabels();

        var cbPreview = dlg.add('checkbox', undefined, L('labelPreview'));
        cbPreview.value = (ss.preview !== undefined) ? !!ss.preview : true;

        function parsePercent() {
            var v = Number(et.text);
            if (isNaN(v) || v <= 0) return null;
            return v;
        }

        function parseStrokeWidth() {
            var vUnit = Number(etStroke.text);
            if (isNaN(vUnit) || vUnit <= 0) return null;
            vUnit = Math.round(vUnit * 10) / 10;
            var vPt = unitToPt(vUnit, "strokeUnits");
            if (isNaN(vPt) || vPt <= 0) return null;
            vPt = Math.round(vPt * 10) / 10;
            return vPt;
        }

        function parseCornerRadius() {
            var vUnit = Number(etCorner.text);
            if (isNaN(vUnit) || vUnit < 0) return null;
            vUnit = Math.round(vUnit * 10) / 10;
            var vPt = unitToPt(vUnit, "rulerType");
            if (isNaN(vPt) || vPt < 0) return null;
            vPt = Math.round(vPt * 10) / 10;
            return vPt;
        }

        function getBalanceMode() {
            if (rbPinLeft.value) return 'left';
            if (rbPinRight.value) return 'right';
            return 'none';
        }

        function parseWidthPercent() {
            var vUnit = Number(etWidth.text);
            if (isNaN(vUnit) || vUnit < 0) vUnit = 0;
            vUnit = Math.round(vUnit * 10) / 10;

            try { slWidth.value = vUnit; } catch (e0) { }
            try { etWidth.text = String(vUnit); } catch (e1) { }

            var vPt = unitToPt(vUnit, "rulerType");
            if (isNaN(vPt) || vPt < 0) vPt = 0;
            vPt = Math.round(vPt * 10) / 10;
            return vPt;
        }

        function applyPreview() {
            var v = parsePercent();
            if (v === null) { clearPreviewFn(); return; }

            var sw = parseStrokeWidth();
            if (sw === null) { clearPreviewFn(); return; }

            var cr = parseCornerRadius();
            if (cr === null) { clearPreviewFn(); return; }

            var bm = getBalanceMode();
            var wp = parseWidthPercent();

            previewFn(
                v,
                cbPreview.value,
                cbOverallFrame.value,
                cbDivider.value,
                cbFillLeft.value,
                cbFillRight.value,
                sw,
                cr,
                bm,
                wp,
                getDirectionMode()
            );
        }

        function persistCurrentUIToSession() {
            try {
                var s = getSessionSettings();
                s.percent = Number(et.text);
                if (isNaN(s.percent) || s.percent <= 0) s.percent = 200;

                s.preview = !!cbPreview.value;
                s.overallFrame = !!cbOverallFrame.value;
                s.divider = !!cbDivider.value;
                s.fillLeft = !!cbFillLeft.value;
                s.fillRight = !!cbFillRight.value;

                s.strokeUnit = String(etStroke.text);
                s.cornerUnit = String(etCorner.text);

                s.balanceMode = getBalanceMode();
                s.widthUnit = Number(etWidth.text);
                if (isNaN(s.widthUnit) || s.widthUnit < 0) s.widthUnit = 0;
                s.widthUnit = Math.round(s.widthUnit * 10) / 10;

                saveSessionSettings(s);
            } catch (e) { }
        }

        (function () {
            var prevOnShow = dlg.onShow;
            dlg.onShow = function () {
                if (typeof prevOnShow === "function") {
                    try { prevOnShow(); } catch (ePrevShow) { }
                }
                try { updateWidthSliderMaxBySelection(AUTO_DIRECTION_MODE); } catch (eMaxInit) { }
                try { updateDirectionUILabels(); } catch (eDirInit) { }
                try { applyPreview(); } catch (eInitPrev) { }
            };
        })();

        et.addEventListener('changing', function () { applyPreview(); });
        etStroke.addEventListener('changing', function () { applyPreview(); });
        etCorner.addEventListener('changing', function () { applyPreview(); });

        cbPreview.onClick = function () { applyPreview(); };
        cbOverallFrame.onClick = function () { updateStrokeWidthEnabled(); updateCornerEnabled(); applyPreview(); };
        cbDivider.onClick = function () { updateStrokeWidthEnabled(); applyPreview(); };
        cbFillLeft.onClick = function () { updateCornerEnabled(); applyPreview(); };
        cbFillRight.onClick = function () { updateCornerEnabled(); applyPreview(); };

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
            // プレビューはヒストリーを汚さないように一括Undoしてから確定描画する
            try { PreviewHistory.undo(); } catch (ePHU1) { }
            // Undo により一時アウトライン等が復活することがあるため、TEMPマーカーを掃除
            try { removeMarkedTempItems(doc); } catch (eTmpU1) { }

            // NOTE: clearPreviewFn() は last* 状態を初期化してしまうため、OK時には呼ばない。
            // プレビュー生成物は undo 済みで存在しないので、参照だけクリアしておく。
            previewRect1 = null;
            previewRect2 = null;
            previewFrame = null;
            previewDivider = null;
            previewApplied = false;
            lastPreviewPercent = null;

            persistCurrentUIToSession();
            dlg.close(1);
        };

        cancelBtn.onClick = function () {
            try { PreviewHistory.undo(); } catch (ePHU2) { }
            // Undo により一時アウトライン等が復活することがあるため、TEMPマーカーを掃除
            try { removeMarkedTempItems(doc); } catch (eTmpU2) { }
            clearPreviewFn();
            persistCurrentUIToSession();
            dlg.close(0);
        };

        dlg.defaultElement = okBtn;
        dlg.cancelElement = cancelBtn;

        dlg.onClose = function () {
            try { PreviewHistory.undo(); } catch (ePHU3) { }
            // Undo により一時アウトライン等が復活することがあるため、TEMPマーカーを掃除
            try { removeMarkedTempItems(doc); } catch (eTmpU3) { }
            persistCurrentUIToSession();
            return true;
        };

        var result = dlg.show();
        if (result !== 1) return null;

        var percent = parsePercent();
        if (percent === null) return null;

        return percent;
    }

    var DEFAULT_HEIGHT_PERCENT = (AUTO_DIRECTION_MODE === 'vertical') ? 130 : 200;

    function previewFn(percent, enabled, addOverallFrame, addDivider, addFillLeft, addFillRight, strokeWidth, cornerRadius, balanceMode, widthPercent, directionMode) {
        clearPreviewRects();
        if (!enabled) return;

        try {
            var sr = percent / 100;
            var result = drawBackgroundRects(itemA, itemB, sr, addOverallFrame, addDivider, addFillLeft, addFillRight, strokeWidth, cornerRadius, balanceMode, widthPercent, directionMode);

            previewRect1 = result.rect1;
            previewRect2 = result.rect2;
            previewFrame = result.frame;
            previewDivider = result.divider;

            if (previewRect1) markPreview(previewRect1);
            if (previewRect2) markPreview(previewRect2);
            if (previewFrame) markPreview(previewFrame);
            if (previewDivider) markPreview(previewDivider);

            try { PreviewHistory.bump(); } catch (ePH1) { }

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
            lastDirectionMode = directionMode || 'horizontal';

            safeRedraw();
        } catch (ePrev) {
            clearPreviewRects();
            throw ePrev;
        }
    }

    var heightPercent = showHeightDialog(DEFAULT_HEIGHT_PERCENT, previewFn, clearPreviewRects);
    if (heightPercent === null) return;

    var sizeRatio = heightPercent / 100;

    // --- 確定処理 ---
    // ［OK］時は PreviewHistory.undo() によりプレビュー生成物は消えているため、
    // 必ずここで最終描画を行い「OK直前の見た目」で確定する。
    var finalResult = drawBackgroundRects(
        itemA, itemB,
        sizeRatio,
        lastOverallFrame, lastDivider,
        lastFillLeft, lastFillRight,
        lastStrokeWidth, lastCornerRadius,
        lastBalanceMode, lastWidthPercent,
        lastDirectionMode
    );
    previewRect1 = finalResult.rect1;
    previewRect2 = finalResult.rect2;
    previewFrame = finalResult.frame;
    previewDivider = finalResult.divider;

    // 確定描画はプレビュー扱いにしない
    if (previewRect1) unmarkPreview(previewRect1);
    if (previewRect2) unmarkPreview(previewRect2);
    if (previewFrame) unmarkPreview(previewFrame);
    if (previewDivider) unmarkPreview(previewDivider);

    doc.selection = null;

    // ==============================
    // Round Any Corner / 角丸アルゴリズム（選択アンカーのみ丸める）
    // Based on: Hiroyuki Sato (MIT) https://github.com/shspage
    // ==============================

    function roundAnyCorner(s, conf) {
        var rr = conf.rr;

        var p, op, pnts;
        var skipList, adjRdirAtEnd, redrawFlg;
        var i, nxi, pvi, q, d, ds, r, g, t, qb;
        var anc1, ldir1, rdir1, anc2, ldir2, rdir2;

        var hanLen = 4 * (Math.sqrt(2) - 1) / 3;
        var ptyp = PointType.SMOOTH;

        for (var j = 0; j < s.length; j++) {
            p = s[j].pathPoints;
            if (readjustAnchors(p) < 2) continue;
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
                if (arrEq(q[0], q[1]) && arrEq(q[2], q[3])) {
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
                        if (r < rr) {
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
                } else {
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
                    } else {
                        if (r < rr) {
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
                        } else {
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
                for (i = p.length - 1; i > 0; i--) p[i].remove();

                for (i = 0; i < pnts.length; i++) {
                    var pt = i > 0 ? p.add() : p[0];
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
    }

    function getPnt(pt, rad, len) {
        return [pt[0] + Math.cos(rad) * len,
        pt[1] + Math.sin(rad) * len];
    }

    function defHan(t, q, n) {
        return [t * (t * (q[n][0] - 2 * q[n + 1][0] + q[n + 2][0]) + 2 * (q[n + 1][0] - q[n][0])) + q[n][0],
        t * (t * (q[n][1] - 2 * q[n + 1][1] + q[n + 2][1]) + 2 * (q[n + 1][1] - q[n][1])) + q[n][1]];
    }

    function bezier(q, t) {
        var u = 1 - t;
        return [u * u * u * q[0][0] + 3 * u * t * (u * q[1][0] + t * q[2][0]) + t * t * t * q[3][0],
        u * u * u * q[0][1] + 3 * u * t * (u * q[1][1] + t * q[2][1]) + t * t * t * q[3][1]];
    }

    function adjHan(anc, dir, m) {
        return [anc[0] + (dir[0] - anc[0]) * m,
        anc[1] + (dir[1] - anc[1]) * m];
    }

    function isCorner(p, idx) {
        var pnt0 = getAnglePnt(p, idx, -1);
        var pnt1 = getAnglePnt(p, idx, 1);
        if (!pnt0 || !pnt1) return false;
        if (pnt0.length < 1 || pnt1.length < 1) return false;
        var rad = getRad2(pnt0, p[idx].anchor, pnt1, true);
        if (rad > Math.PI - 0.1) return false;
        return true;
    }

    function getAnglePnt(p, idx1, dir) {
        if (!dir) dir = -1;
        var idx2 = parseIdx(p, idx1 + dir);
        if (idx2 < 0) return null;
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

    function arrEq(arr1, arr2) {
        for (var i = 0; i < arr1.length; i++) {
            if (arr1[i] != arr2[i]) return false;
        }
        return true;
    }

    function dist(p1, p2) {
        return Math.sqrt(Math.pow(p1[0] - p2[0], 2)
            + Math.pow(p1[1] - p2[1], 2));
    }

    function dist2(p1, p2) {
        return Math.pow(p1[0] - p2[0], 2)
            + Math.pow(p1[1] - p2[1], 2);
    }

    function getRad(p1, p2) {
        return Math.atan2(p2[1] - p1[1],
            p2[0] - p1[0]);
    }

    function getRad2(p1, o, p2) {
        var v1 = normalize(p1, o);
        var v2 = normalize(p2, o);
        return Math.acos(v1[0] * v2[0] + v1[1] * v2[1]);
    }

    function normalize(p, o) {
        var d = dist(p, o);
        return d == 0 ? [0, 0] : [(p[0] - o[0]) / d,
        (p[1] - o[1]) / d];
    }

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

    function getLength(k, t) {
        var h = t / 128;
        var hh = h * 2;
        var fc = function (t, k) {
            return Math.sqrt(t * (t * (t * (t * k[0] + k[1]) + k[2]) + k[3]) + k[4]) || 0;
        };
        var total = (fc(0, k) - fc(t, k)) / 2;
        for (var i = h; i < t; i += hh) total += 2 * fc(i, k) + fc(i + h, k);
        return total * hh;
    }

    function readjustAnchors(p) {
        var minDist = 0.0025;
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

    function parseIdx(p, n) {
        var len = p.length;
        if (p.parent.closed) {
            return n >= 0 ? n % len : len - Math.abs(n % len);
        } else {
            return (n < 0 || n > len - 1) ? -1 : n;
        }
    }

    function getDat(p) {
        with (p) return [anchor, rightDirection, leftDirection, pointType];
    }

    function isSelected(p) {
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
            if (sel > 0 && sel < 4) continue;

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