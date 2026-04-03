#target illustrator
#targetengine "SplitBackgroundForTwoEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：
SplitBackgroundForTwo

### 更新日：
20260310

### 概要：
1つのオブジェクト（テキスト、パス、グループなど）を選択して実行すると、そのオブジェクトの外接矩形を左右または上下に2分割し、背面に2色の背景を作成します。
プレビューは専用レイヤーに描画して差し替え、［OK］時に図形や線が二重に作成されないようにします。

［分割方法］で「左右／上下」を選択できます。
［バランス］パネルでは、左・右（または上・下）の幅と比率を数値入力およびスライダーで調整できます。

［線］では外枠と区切り線の有無、線幅、線色を指定できます。
［オプション］では角丸やピル形状を指定できます。
カラー指定は RGB / CMYK / グレーに対応したカラーピッカーから行えます。
プリセットの保存・呼び出し・書き出しにも対応しています。

テキストのサイドベアリング等によるズレを減らすため、
一時的にテキストをアウトライン化して外接矩形を計算し、計算後すぐに一時生成物を削除します。
その上で背景長方形を選択オブジェクトの背面に配置し、ダイアログ内でリアルタイムにプレビュー表示します（元のオブジェクトは変更しません）。

［幅］入力欄では、↑↓キーで±1、Shift+↑↓で±10（10刻みスナップ）、Option+↑↓で±0.1 の増減が可能です。

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
- v2.7 (20260228) : プレビュー更新時に直前のプレビューを1回Undoして差し替える方式に変更し、［OK］時に図形や線が二重に生成される問題を修正。
- v2.8 (20260228) : プレビューのロールバックをUndo依存から、プレビュー印（PREVIEW_NOTE）付きオブジェクトの直接削除方式に変更し、線幅プレビューが更新されない／線がダブる問題を修正。
- v2.9 (20260228) : プレビュー専用レイヤーに描画してレイヤー単位で掃除する方式に変更し、プレビュー残骸やダブりをさらに抑止。
- v2.9.1 (20260310) : ローカライズ漏れのあった固定文言・ボタン名・パネル名・アラート文を LABELS 経由へ統一し、英語環境でもUI全体が切り替わるよう改善。
*/

// --- Version / バージョン ---
var SCRIPT_VERSION = "v2.9.1";

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
        ja: "オブジェクトの分割背景を作成",
        en: "Create Split Background"
    },

    colorPickerTitle: {
        ja: "カラー選択",
        en: "Color Picker"
    },
    labelWhite: {
        ja: "白",
        en: "White"
    },
    labelBlack: {
        ja: "黒",
        en: "Black"
    },
    labelCustom: {
        ja: "カスタム",
        en: "Custom"
    },
    labelRGB: {
        ja: "RGB",
        en: "RGB"
    },
    labelCMYK: {
        ja: "CMYK",
        en: "CMYK"
    },
    labelGray: {
        ja: "グレー",
        en: "Gray"
    },
    panelPreset: {
        ja: "プリセット",
        en: "Preset"
    },
    labelSave: {
        ja: "保存",
        en: "Save"
    },
    labelExport: {
        ja: "書き出し",
        en: "Export"
    },
    presetSaveTitle: {
        ja: "プリセット保存",
        en: "Save Preset"
    },
    presetNamePrompt: {
        ja: "プリセット名を入力:",
        en: "Enter preset name:"
    },
    panelSplitDirection: {
        ja: "分割方法",
        en: "Split Direction"
    },
    labelSplitLR: {
        ja: "左右",
        en: "Left/Right"
    },
    labelSplitTB: {
        ja: "上下",
        en: "Top/Bottom"
    },
    panelFill: {
        ja: "塗り",
        en: "Fill"
    },
    labelColor: {
        ja: "カラー",
        en: "Color"
    },
    panelOptions: {
        ja: "オプション",
        en: "Options"
    },
    labelPillShape: {
        ja: "ピル形状",
        en: "Pill shape"
    },
    labelTop: {
        ja: "上",
        en: "Top"
    },
    labelBottom: {
        ja: "下",
        en: "Bottom"
    },
    alertExportedSettings: {
        ja: "現在の設定を書き出しました:",
        en: "Exported current settings:"
    },
    alertExportFailed: {
        ja: "ファイルの書き出しに失敗しました。",
        en: "Failed to export the file."
    },

    alertOpenDoc: {
        ja: "ドキュメントを開いてください。",
        en: "Please open a document."
    },
    alertSelectOneItem: {
        ja: "1つのオブジェクトを選択してください。",
        en: "Please select one object."
    },

    labelPreview: {
        ja: "プレビュー",
        en: "Preview"
    },
    labelOverallFrame: {
        ja: "外枠",
        en: "Outer frame"
    },
    labelDivider: {
        ja: "区切り線",
        en: "Divider"
    },
    labelFillLeft: {
        ja: "左",
        en: "Left"
    },
    labelFillRight: {
        ja: "右",
        en: "Right"
    },
    panelLine: {
        ja: "線",
        en: "Stroke"
    },
    // --- 固定パネルラベル追加 ---
    panelFixed: {
        ja: "バランス",
        en: "Balance"
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
                preview: true,
                overallFrame: false,
                divider: false,
                fillLeft: true,
                fillRight: true,
                strokeUnit: formatUnitValue(ptToUnit(1, "strokeUnits")),
                cornerUnit: formatUnitValue(ptToUnit(0, "rulerType")),
                widthUnit: 0
            };
        }
        return $.global[SETTINGS_KEY];
    } catch (e) {
        return {
            preview: true,
            overallFrame: false,
            divider: false,
            fillLeft: true,
            fillRight: true,
            strokeUnit: formatUnitValue(ptToUnit(1, "strokeUnits")),
            cornerUnit: formatUnitValue(ptToUnit(0, "rulerType")),
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

    if (!sel || sel.length !== 1) {
        alert(L("alertSelectOneItem"));
        return;
    }

    var item = sel[0];

    // --- Temporary marker (cleanup) / 一時生成物マーカー（後片付け） ---
    var TEMP_NOTE = "__SplitBackgroundForTwo_TEMP__";
    var TEMP_NAME_PREFIX = "__SplitBackgroundForTwo_TEMP__";

    // --- Preview marker (cleanup) / プレビューマーカー（後片付け） ---
    var PREVIEW_NOTE = "__SplitBackgroundForTwo_PREVIEW__";

    // --- Preview layer (robust preview container) / プレビューレイヤー（堅牢なプレビュー用コンテナ） ---
    var PREVIEW_LAYER_NAME = "__SplitBackgroundForTwo__PreviewLayer__";

    function findLayerByName(doc, name) {
        try {
            for (var i = 0; i < doc.layers.length; i++) {
                if (doc.layers[i].name === name) return doc.layers[i];
            }
        } catch (e) { }
        return null;
    }

    function ensurePreviewLayer(doc, layerA, layerB) {
        var lyr = findLayerByName(doc, PREVIEW_LAYER_NAME);
        try {
            if (!lyr) {
                lyr = doc.layers.add();
                lyr.name = PREVIEW_LAYER_NAME;
            }

            // 2レイヤーのうち“背面側”よりさらに背面へ
            var backmost = getBackmostLayer(doc, layerA, layerB) || layerA || layerB;
            if (backmost && backmost.isValid) {
                try { lyr.move(backmost, ElementPlacement.PLACEAFTER); } catch (eMv) { }
            }

            try { lyr.visible = true; } catch (eV) { }
            try { lyr.locked = false; } catch (eL) { }
            try { lyr.printable = true; } catch (eP) { }
            try { lyr.template = false; } catch (eT) { }
        } catch (e2) { }
        return lyr;
    }

    function clearLayerItems(layer) {
        if (!layer) return;
        try { layer.locked = false; } catch (e0) { }
        try { layer.visible = true; } catch (e1) { }
        try {
            for (var i = layer.pageItems.length - 1; i >= 0; i--) {
                var it = layer.pageItems[i];
                try { unlockAndShow(it); } catch (eU) { }
                try { it.remove(); } catch (eR) { }
            }
        } catch (eLoop) { }
    }

    function removePreviewLayerIfExists(doc) {
        var lyr = findLayerByName(doc, PREVIEW_LAYER_NAME);
        if (!lyr) return;
        try {
            clearLayerItems(lyr);
            lyr.remove();
        } catch (e) { }
    }

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
                dec: function () {
                    var n = g.__previewUndoCount | 0;
                    g.__previewUndoCount = (n > 0) ? (n - 1) : 0;
                },
                rollbackOne: function () {
                    var n = g.__previewUndoCount | 0;
                    if (n <= 0) return;
                    try { app.executeMenuCommand('undo'); } catch (e) { }
                    g.__previewUndoCount = (n > 0) ? (n - 1) : 0;
                },
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

    // Remove current preview objects without touching last* state.
    function clearPreviewItemsOnly(doc) {
        // Prefer clearing the dedicated preview layer (most robust)
        try {
            var lyr = findLayerByName(doc, PREVIEW_LAYER_NAME);
            if (lyr && lyr.isValid) {
                clearLayerItems(lyr);
            }
        } catch (eLy) { }

        // Fallback: remove by PREVIEW_NOTE
        try { removeMarkedPreviewItems(doc); } catch (e0) { }

        previewState.rect1 = null;
        previewState.rect2 = null;
        previewState.frame = null;
        previewState.divider = null;
        previewState.applied = false;
        safeRedraw();
    }

    function hasMarkedPreviewItems(doc) {
        try {
            for (var i = doc.pageItems.length - 1; i >= 0; i--) {
                var it = doc.pageItems[i];
                try {
                    if (it.note === PREVIEW_NOTE) return true;
                } catch (eNote) { }
            }
        } catch (eLoop) { }
        return false;
    }

    // Roll back preview safely even if a single preview refresh created multiple undo steps.
    // We undo repeatedly until preview-marked items are gone (with a hard cap), then reset the counter.
    function rollbackPreviewSafely(doc) {
        var max = 30;
        try {
            while (max-- > 0) {
                if (!hasMarkedPreviewItems(doc)) break;
                try { app.executeMenuCommand('undo'); } catch (eU) { break; }
            }
        } catch (e) { }
        // Reset counter because we can no longer trust the exact undo-step count.
        try { PreviewHistory.start(); } catch (ePH) { }
        // Undo により一時アウトライン等が復活することがあるため、TEMPマーカーを掃除
        try { removeMarkedTempItems(doc); } catch (eTmp) { }
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

    // =============================================
    // カラーシステム
    // 値の形式: "RRGGBB" (RGB HEX) または "cmyk:C,M,Y,K" (CMYK 0-100)
    // =============================================

    function isCmykString(s) { return String(s).indexOf('cmyk:') === 0; }

    function parseCmykString(s) {
        var parts = String(s).replace('cmyk:', '').split(',');
        var c = parseFloat(parts[0]) || 0;
        var m = parseFloat(parts[1]) || 0;
        var y = parseFloat(parts[2]) || 0;
        var k = parseFloat(parts[3]) || 0;
        return { c: c, m: m, y: y, k: k };
    }

    function cmykStringFromValues(c, m, y, k) {
        return 'cmyk:' + Math.round(c) + ',' + Math.round(m) + ',' + Math.round(y) + ',' + Math.round(k);
    }

    // HEX→RGBColor
    function hexToRGBColor(hex) {
        hex = String(hex).replace(/^#/, '');
        if (hex.length !== 6) hex = 'CCCCCC';
        var r = parseInt(hex.substring(0, 2), 16);
        var g = parseInt(hex.substring(2, 4), 16);
        var b = parseInt(hex.substring(4, 6), 16);
        if (isNaN(r)) r = 200; if (isNaN(g)) g = 200; if (isNaN(b)) b = 200;
        var col = new RGBColor();
        col.red = r; col.green = g; col.blue = b;
        return col;
    }

    // CMYK値→CMYKColor
    function cmykValuesToColor(c, m, y, k) {
        var col = new CMYKColor();
        col.cyan = c; col.magenta = m; col.yellow = y; col.black = k;
        return col;
    }

    // 文字列からIllustrator Colorオブジェクトを生成
    function parseColorString(s) {
        if (isCmykString(s)) {
            var v = parseCmykString(s);
            return cmykValuesToColor(v.c, v.m, v.y, v.k);
        }
        return hexToRGBColor(s);
    }

    // RGB→HEX
    function rgbToHex(r, g, b) {
        function toHex2(n) {
            var h = Math.round(n).toString(16).toUpperCase();
            return h.length < 2 ? '0' + h : h;
        }
        return toHex2(r) + toHex2(g) + toHex2(b);
    }

    // 簡易 CMYK→RGB 変換（プレビュー表示用、近似値）
    function cmykToRgbApprox(c, m, y, k) {
        var r = 255 * (1 - c / 100) * (1 - k / 100);
        var g = 255 * (1 - m / 100) * (1 - k / 100);
        var b = 255 * (1 - y / 100) * (1 - k / 100);
        return { r: Math.round(r), g: Math.round(g), b: Math.round(b) };
    }

    // 簡易 RGB→CMYK 変換（近似値）
    function rgbToCmykApprox(r, g, b) {
        var rr = r / 255, gg = g / 255, bb = b / 255;
        var k = 1 - Math.max(rr, gg, bb);
        if (k >= 1) return { c: 0, m: 0, y: 0, k: 100 };
        var c = (1 - rr - k) / (1 - k) * 100;
        var m = (1 - gg - k) / (1 - k) * 100;
        var y = (1 - bb - k) / (1 - k) * 100;
        return { c: Math.round(c), m: Math.round(m), y: Math.round(y), k: Math.round(k * 100) };
    }

    // カラー文字列からプレビュー用RGB(0-1)を取得
    function colorStringToRgbNorm(s) {
        if (isCmykString(s)) {
            var v = parseCmykString(s);
            var rgb = cmykToRgbApprox(v.c, v.m, v.y, v.k);
            return [rgb.r / 255, rgb.g / 255, rgb.b / 255];
        }
        var hex = String(s).replace(/^#/, '');
        if (hex.length !== 6) hex = '000000';
        return [
            (parseInt(hex.substring(0, 2), 16) || 0) / 255,
            (parseInt(hex.substring(2, 4), 16) || 0) / 255,
            (parseInt(hex.substring(4, 6), 16) || 0) / 255
        ];
    }

    // カラー文字列の表示用ラベル
    function colorDisplayLabel(s) {
        if (isCmykString(s)) {
            var v = parseCmykString(s);
            return 'C' + Math.round(v.c) + ' M' + Math.round(v.m) + ' Y' + Math.round(v.y) + ' K' + Math.round(v.k);
        }
        return '#' + String(s).replace(/^#/, '').toUpperCase();
    }

    // =============================================
    // ScriptUI カラーピッカーダイアログ（RGB/CMYK対応）
    // =============================================
    var _colorPickerLastPos = null; // ダイアログ位置記憶用

    function showColorPickerDialog(initColorStr) {
        var isCmyk = isCmykString(initColorStr);
        var ir = 0, ig = 0, ib = 0;
        var ic = 0, im = 0, iy = 0, ik = 0;

        if (isCmyk) {
            var cv = parseCmykString(initColorStr);
            ic = cv.c; im = cv.m; iy = cv.y; ik = cv.k;
            var rgbInit = cmykToRgbApprox(ic, im, iy, ik);
            ir = rgbInit.r; ig = rgbInit.g; ib = rgbInit.b;
        } else {
            var hex = String(initColorStr).replace(/^#/, '');
            if (hex.length !== 6) hex = '000000';
            ir = parseInt(hex.substring(0, 2), 16) || 0;
            ig = parseInt(hex.substring(2, 4), 16) || 0;
            ib = parseInt(hex.substring(4, 6), 16) || 0;
            var cmykInit = rgbToCmykApprox(ir, ig, ib);
            ic = cmykInit.c; im = cmykInit.m; iy = cmykInit.y; ik = cmykInit.k;
        }

        var d = new Window('dialog', L('colorPickerTitle'));
        d.orientation = 'column';
        d.alignChildren = ['fill', 'top'];

        // プレビュー（左：初期色、右：現在色）
        var previewRow = d.add('group');
        previewRow.orientation = 'row';
        previewRow.alignment = ['center', 'top'];
        previewRow.spacing = 0;
        var previewOrig = previewRow.add('group');
        previewOrig.preferredSize = [100, 40];
        var previewCur = previewRow.add('group');
        previewCur.preferredSize = [100, 40];

        // 初期色のRGB値を保存
        var initRgb = isCmyk
            ? cmykToRgbApprox(ic, im, iy, ik)
            : { r: ir, g: ig, b: ib };
        previewOrig.onDraw = function () {
            var gr = this.graphics;
            var brush = gr.newBrush(gr.BrushType.SOLID_COLOR, [initRgb.r / 255, initRgb.g / 255, initRgb.b / 255, 1]);
            gr.rectPath(0, 0, this.size[0], this.size[1]);
            gr.fillPath(brush);
        };

        // --- 白/黒/カスタム切替 ---
        var colorTypeRow = d.add('group');
        colorTypeRow.orientation = 'row';
        colorTypeRow.alignment = ['center', 'top'];
        colorTypeRow.alignChildren = ['left', 'center'];
        var rbWhite = colorTypeRow.add('radiobutton', undefined, L('labelWhite'));
        var rbBlack = colorTypeRow.add('radiobutton', undefined, L('labelBlack'));
        var rbCustom = colorTypeRow.add('radiobutton', undefined, L('labelCustom'));

        // 初期値の判定
        var isWhite = (!isCmyk && ir === 255 && ig === 255 && ib === 255)
            || (isCmyk && ic === 0 && im === 0 && iy === 0 && ik === 0);
        var isBlack = (!isCmyk && ir === 0 && ig === 0 && ib === 0)
            || (isCmyk && ic === 0 && im === 0 && iy === 0 && ik === 100);
        if (isWhite) { rbWhite.value = true; }
        else if (isBlack) { rbBlack.value = true; }
        else { rbCustom.value = true; }

        // モード切替
        var modeRow = d.add('group');
        modeRow.orientation = 'row';
        modeRow.alignment = ['center', 'top'];
        modeRow.alignChildren = ['left', 'center'];
        var rbRGB = modeRow.add('radiobutton', undefined, L('labelRGB'));
        var rbCMYK = modeRow.add('radiobutton', undefined, L('labelCMYK'));
        var rbGray = modeRow.add('radiobutton', undefined, L('labelGray'));

        // 初期モード判定（グレー: CMYK で C=M=Y=0, K>0）
        var isGrayInit = isCmyk && ic === 0 && im === 0 && iy === 0 && ik > 0;
        if (isGrayInit) {
            rbGray.value = true;
        } else {
            rbRGB.value = !isCmyk;
            rbCMYK.value = isCmyk;
        }

        // --- RGB パネル ---
        var rgbPanel = d.add('panel');
        rgbPanel.text = '';
        rgbPanel.orientation = 'column';
        rgbPanel.alignChildren = ['fill', 'top'];

        function makeSliderRow(parent, label, initVal, maxVal) {
            var row = parent.add('group');
            row.orientation = 'row';
            row.alignChildren = ['left', 'center'];
            var st = row.add('statictext', undefined, label);
            st.justify = 'center';
            st.preferredSize = [20, -1];
            var sl = row.add('slider', undefined, initVal, 0, maxVal);
            sl.preferredSize = [120, 20];
            var et = row.add('edittext', undefined, String(Math.round(initVal)));
            et.characters = 3;
            sl.onChanging = function () {
                var v = sl.value;
                var keyboard = ScriptUI.environment.keyboardState;
                if (keyboard.shiftKey) {
                    var step = maxVal * 0.1;
                    v = Math.round(v / step) * step;
                } else if (keyboard.altKey) {
                    var stepAlt = (maxVal === 255) ? 16 : maxVal * 0.05;
                    v = Math.round(v / stepAlt) * stepAlt;
                    if (maxVal === 255 && v > 255) v = 255;
                }
                v = Math.round(v);
                if (v < 0) v = 0;
                if (v > maxVal) v = maxVal;
                sl.value = v;
                et.text = String(v);
            };
            et.addEventListener('changing', function () {
                var v = parseFloat(et.text);
                if (isNaN(v)) v = 0;
                if (v < 0) v = 0;
                if (v > maxVal) v = maxVal;
                sl.value = v;
            });
            var _onKeyChange = null;
            et.addEventListener('keydown', function (event) {
                if (event.keyName !== 'Up' && event.keyName !== 'Down') return;
                var v = parseFloat(et.text);
                if (isNaN(v)) v = 0;
                var keyboard = ScriptUI.environment.keyboardState;
                var delta = keyboard.shiftKey ? 10 : 1;
                if (event.keyName === 'Up') v += delta;
                else v -= delta;
                if (v < 0) v = 0;
                if (v > maxVal) v = maxVal;
                v = Math.round(v);
                et.text = String(v);
                sl.value = v;
                if (typeof _onKeyChange === 'function') _onKeyChange();
                event.preventDefault();
            });
            var result = { slider: sl, edit: et, row: row };
            result.setOnKeyChange = function (fn) { _onKeyChange = fn; };
            return result;
        }

        var rRow = makeSliderRow(rgbPanel, 'R', ir, 255);
        var gRow = makeSliderRow(rgbPanel, 'G', ig, 255);
        var bRow = makeSliderRow(rgbPanel, 'B', ib, 255);
        var slR = rRow.slider, slG = gRow.slider, slB = bRow.slider;

        // --- HEX 入力（ダイアログ中央） ---
        var hexRow = d.add('group');
        hexRow.orientation = 'row';
        hexRow.alignment = ['center', 'top'];
        hexRow.alignChildren = ['left', 'center'];
        hexRow.add('statictext', undefined, '#');
        var etHex = hexRow.add('edittext', undefined, rgbToHex(ir, ig, ib));
        etHex.characters = 6;

        // --- CMYK パネル ---
        var cmykPanel = d.add('panel');
        cmykPanel.text = '';
        cmykPanel.orientation = 'column';
        cmykPanel.alignChildren = ['fill', 'top'];

        var cRow = makeSliderRow(cmykPanel, 'C', ic, 100);
        var mRow = makeSliderRow(cmykPanel, 'M', im, 100);
        var yRow = makeSliderRow(cmykPanel, 'Y', iy, 100);
        var kRow = makeSliderRow(cmykPanel, 'K', ik, 100);
        var slC = cRow.slider, slM = mRow.slider, slY = yRow.slider, slK = kRow.slider;

        // プレビュー描画（常にRGB近似で表示）
        function getPreviewRGB() {
            if (rbWhite.value) return { r: 255, g: 255, b: 255 };
            if (rbBlack.value) return { r: 0, g: 0, b: 0 };
            if (rbGray.value) {
                return cmykToRgbApprox(0, 0, 0, slK.value);
            }
            if (rbCMYK.value) {
                return cmykToRgbApprox(slC.value, slM.value, slY.value, slK.value);
            }
            return { r: Math.round(slR.value), g: Math.round(slG.value), b: Math.round(slB.value) };
        }

        previewCur.onDraw = function () {
            var gr = this.graphics;
            var rgb = getPreviewRGB();
            var brush = gr.newBrush(gr.BrushType.SOLID_COLOR, [rgb.r / 255, rgb.g / 255, rgb.b / 255, 1]);
            gr.rectPath(0, 0, this.size[0], this.size[1]);
            gr.fillPath(brush);
        };

        function refreshPreview() {
            try {
                previewCur.hide();
                previewCur.show();
            } catch (e) { }
        }

        // モード切替の表示制御（ディム表示）
        function updateModeVisibility() {
            var custom = rbCustom.value;
            var gray = rbGray.value;
            modeRow.enabled = custom;
            rgbPanel.enabled = custom && rbRGB.value;
            hexRow.enabled = custom && rbRGB.value;
            cmykPanel.enabled = custom && (rbCMYK.value || gray);
            // グレーモード: Kのみアクティブ、C/M/Yはディム
            cRow.row.enabled = custom && rbCMYK.value;
            mRow.row.enabled = custom && rbCMYK.value;
            yRow.row.enabled = custom && rbCMYK.value;
            kRow.row.enabled = custom && (rbCMYK.value || gray);
        }

        // RGB→CMYK同期
        function syncCmykFromRgb() {
            var cv = rgbToCmykApprox(Math.round(slR.value), Math.round(slG.value), Math.round(slB.value));
            slC.value = cv.c; cRow.edit.text = String(cv.c);
            slM.value = cv.m; mRow.edit.text = String(cv.m);
            slY.value = cv.y; yRow.edit.text = String(cv.y);
            slK.value = cv.k; kRow.edit.text = String(cv.k);
        }

        // CMYK→RGB同期
        function syncRgbFromCmyk() {
            var rgb = cmykToRgbApprox(slC.value, slM.value, slY.value, slK.value);
            slR.value = rgb.r; rRow.edit.text = String(rgb.r);
            slG.value = rgb.g; gRow.edit.text = String(rgb.g);
            slB.value = rgb.b; bRow.edit.text = String(rgb.b);
            etHex.text = rgbToHex(rgb.r, rgb.g, rgb.b);
        }

        // HEX同期
        function syncHex() {
            etHex.text = rgbToHex(Math.round(slR.value), Math.round(slG.value), Math.round(slB.value));
        }

        // RGBスライダーのイベント拡張
        var origR = slR.onChanging, origG = slG.onChanging, origB = slB.onChanging;
        slR.onChanging = function () { origR.call(slR); syncHex(); refreshPreview(); };
        slG.onChanging = function () { origG.call(slG); syncHex(); refreshPreview(); };
        slB.onChanging = function () { origB.call(slB); syncHex(); refreshPreview(); };

        // CMYKスライダーのイベント拡張
        var origC = slC.onChanging, origM = slM.onChanging, origY = slY.onChanging, origK = slK.onChanging;
        slC.onChanging = function () { origC.call(slC); refreshPreview(); };
        slM.onChanging = function () { origM.call(slM); refreshPreview(); };
        slY.onChanging = function () { origY.call(slY); refreshPreview(); };
        slK.onChanging = function () { origK.call(slK); refreshPreview(); };

        // edittextの↑↓キー変更時コールバックを設定
        rRow.setOnKeyChange(function () { syncHex(); refreshPreview(); });
        gRow.setOnKeyChange(function () { syncHex(); refreshPreview(); });
        bRow.setOnKeyChange(function () { syncHex(); refreshPreview(); });
        cRow.setOnKeyChange(function () { refreshPreview(); });
        mRow.setOnKeyChange(function () { refreshPreview(); });
        yRow.setOnKeyChange(function () { refreshPreview(); });
        kRow.setOnKeyChange(function () { refreshPreview(); });

        // HEX入力イベント
        etHex.addEventListener('changing', function () {
            var h = etHex.text.replace(/^#/, '');
            if (h.length === 6) {
                var hr = parseInt(h.substring(0, 2), 16);
                var hg = parseInt(h.substring(2, 4), 16);
                var hb = parseInt(h.substring(4, 6), 16);
                if (!isNaN(hr)) { slR.value = hr; rRow.edit.text = String(hr); }
                if (!isNaN(hg)) { slG.value = hg; gRow.edit.text = String(hg); }
                if (!isNaN(hb)) { slB.value = hb; bRow.edit.text = String(hb); }
                refreshPreview();
            }
        });

        // 白/黒/カスタム切替
        rbWhite.onClick = function () { updateModeVisibility(); refreshPreview(); };
        rbBlack.onClick = function () { updateModeVisibility(); refreshPreview(); };
        rbCustom.onClick = function () { updateModeVisibility(); refreshPreview(); };

        // モード切替時に相互変換
        rbRGB.onClick = function () {
            syncRgbFromCmyk();
            updateModeVisibility();
            refreshPreview();
        };
        rbCMYK.onClick = function () {
            syncCmykFromRgb();
            updateModeVisibility();
            refreshPreview();
        };
        rbGray.onClick = function () {
            slC.value = 0; cRow.edit.text = '0';
            slM.value = 0; mRow.edit.text = '0';
            slY.value = 0; yRow.edit.text = '0';
            updateModeVisibility();
            refreshPreview();
        };

        var btns = d.add('group');
        btns.orientation = 'row';
        btns.alignment = ['center', 'center'];
        btns.margins = [0, 10, 0, 0];
        btns.add('button', undefined, L('labelCancel'), { name: 'cancel' });
        btns.add('button', undefined, L('labelOK'), { name: 'ok' });

        updateModeVisibility();

        // 前回位置を復元
        if (_colorPickerLastPos) {
            d.onShow = function () {
                try { d.location = _colorPickerLastPos; } catch (e) { }
            };
        }

        var result = d.show();

        // 閉じる前に位置を保存
        try { _colorPickerLastPos = [d.location[0], d.location[1]]; } catch (e) { }

        if (result === 1) {
            if (rbWhite.value) return 'FFFFFF';
            if (rbBlack.value) return '000000';
            if (rbGray.value) {
                return cmykStringFromValues(0, 0, 0, Math.round(slK.value));
            }
            if (rbCMYK.value) {
                return cmykStringFromValues(Math.round(slC.value), Math.round(slM.value), Math.round(slY.value), Math.round(slK.value));
            }
            return rgbToHex(Math.round(slR.value), Math.round(slG.value), Math.round(slB.value));
        }
        return null;
    }

    // =============================================
    // カラースウォッチ
    // =============================================
    function createColorSwatch(parent, colorField, onChangeCallback) {
        var swatch = parent.add('image', undefined, undefined, { name: 'swatch' });
        swatch.preferredSize = [20, 20];

        swatch.onDraw = function () {
            var g = this.graphics;
            var rgb = colorStringToRgbNorm(colorField.text);
            var brush = g.newBrush(g.BrushType.SOLID_COLOR, [rgb[0], rgb[1], rgb[2], 1]);
            g.rectPath(0, 0, this.size[0], this.size[1]);
            g.fillPath(brush);
            var borderPen = g.newPen(g.PenType.SOLID_COLOR, [0, 0, 0, 1], 1);
            g.rectPath(0, 0, this.size[0], this.size[1]);
            g.strokePath(borderPen);
        };

        swatch.addEventListener('click', function () {
            try {
                var result = showColorPickerDialog(colorField.text);
                if (result !== null) {
                    colorField.text = result;
                    try { swatch.notify('onDraw'); } catch (eN) { }
                    if (typeof onChangeCallback === 'function') onChangeCallback();
                }
            } catch (eCP) { }
        });

        colorField.addEventListener('changing', function () {
            try { swatch.notify('onDraw'); } catch (eN) { }
        });

        return swatch;
    }

    // デフォルト色
    var colorLeft = parseColorString('DCDCDC');
    var colorRight = parseColorString('808080');
    var colorStroke = parseColorString('000000');

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
            idx.sort(function (x, y) { return pp[y].anchor[1] - pp[x].anchor[1]; });
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
            idx.sort(function (x, y) { return pp[x].anchor[1] - pp[y].anchor[1]; });
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

    var previewState = {
        rect1: null,
        rect2: null,
        frame: null,
        divider: null,
        applied: false,
        last: {
            overallFrame: false,
            divider: false,
            fillLeft: true,
            fillRight: true,
            strokeWidth: 1,
            cornerRadius: 0,
            widthOffset: 0,      // Offset in pt from center (negative=left, positive=right)
            splitDirection: 'lr' // 'lr' or 'tb'
        }
    };

    function clearPreviewRects() {
        // NOTE: v2.7 以降は通常のプレビュー更新では使わず、例外時の掃除専用
        try {
            if (previewState.rect1 && previewState.rect1.isValid) {
                unlockAndShow(previewState.rect1);
                try { previewState.rect1.remove(); } catch (e1a) { }
            }
        } catch (e1) { }
        try {
            if (previewState.rect2 && previewState.rect2.isValid) {
                unlockAndShow(previewState.rect2);
                try { previewState.rect2.remove(); } catch (e2a) { }
            }
        } catch (e2) { }
        try {
            if (previewState.frame && previewState.frame.isValid) {
                unlockAndShow(previewState.frame);
                try { previewState.frame.remove(); } catch (e3a) { }
            }
        } catch (e3) { }
        try {
            if (previewState.divider && previewState.divider.isValid) {
                unlockAndShow(previewState.divider);
                try { previewState.divider.remove(); } catch (e4a) { }
            }
        } catch (e4) { }

        removeMarkedPreviewItems(doc);

        previewState.rect1 = null;
        previewState.rect2 = null;
        previewState.frame = null;
        previewState.divider = null;
        previewState.applied = false;
        previewState.last.overallFrame = false;
        previewState.last.divider = false;
        previewState.last.fillLeft = true;
        previewState.last.fillRight = true;
        previewState.last.strokeWidth = 1;
        previewState.last.cornerRadius = 0;
        previewState.last.widthOffset = 0;
        previewState.last.splitDirection = 'lr';
        safeRedraw();
    }

    /**
     * 単一オブジェクトの外接矩形を分割し、背景長方形を作成。
     * splitDirection: 'lr'=左右, 'tb'=上下
     * offsetPt で中央からの分割位置オフセットを調整。
     */
    function drawBackgroundRects(targetItem, addOverallFrame, addDivider, addFillLeft, addFillRight, strokeWidthValue, cornerRadius, offsetPt, splitDirection) {
        // --- 一時グループ（アウトライン生成用）---
        var originalActiveLayer = doc.activeLayer;

        var hostLayer = null;
        try { hostLayer = targetItem.layer; } catch (eHL) { hostLayer = null; }
        if (!hostLayer) {
            try { hostLayer = doc.activeLayer; } catch (eHL2) { hostLayer = null; }
        }
        if (!hostLayer) hostLayer = doc.layers[0];

        unlockAndShow(hostLayer);
        var tempGroup = hostLayer.groupItems.add();

        markTemp(tempGroup);
        try { tempGroup.opacity = 0; } catch (eOp) { }
        tempGroup.name = TEMP_NAME_PREFIX + "__temp_outline_group__";

        var isVertical = (splitDirection === 'tb');
        var objLeft, objTop, objRight, objBottom, objWidth, objHeight;
        var splitX, splitY;

        try {
            var outlineInfo = getBoundsInfo(targetItem, tempGroup);
            var b = outlineInfo.bounds;

            unlockAndShow(outlineInfo.outline);
            unlockAndShow(outlineInfo.duplicateItem);
            try { if (outlineInfo.outline && outlineInfo.outline.isValid) outlineInfo.outline.remove(); } catch (eA) { }
            try { if (outlineInfo.duplicateItem && outlineInfo.duplicateItem.isValid) outlineInfo.duplicateItem.remove(); } catch (eB) { }
            removeMarkedTempItems(doc);

            function clamp(v, mn, mx) {
                if (v < mn) return mn;
                if (v > mx) return mx;
                return v;
            }

            var oPt = Number(offsetPt);
            if (isNaN(oPt)) oPt = 0;

            objLeft = b[0]; objTop = b[1]; objRight = b[2]; objBottom = b[3];
            objWidth = objRight - objLeft;
            objHeight = objTop - objBottom;

            if (isVertical) {
                // 上下分割: オフセット正=下方向（splitYが下がる）
                splitY = clamp(objTop - objHeight / 2 - oPt, objBottom, objTop);
            } else {
                // 左右分割
                splitX = clamp(objLeft + objWidth / 2 + oPt, objLeft, objRight);
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

        // --- ここから描画 ---
        var itemLayer = null;
        try { itemLayer = targetItem.layer; } catch (eIL) { itemLayer = null; }

        var targetLayer = null;
        try {
            targetLayer = findLayerByName(doc, PREVIEW_LAYER_NAME);
            if (!targetLayer || !targetLayer.isValid) {
                targetLayer = itemLayer;
            }
        } catch (eTL) {
            targetLayer = itemLayer;
        }

        var rect1 = null;
        var rect2 = null;
        var pathItems = targetLayer ? targetLayer.pathItems : doc.pathItems;

        if (isVertical) {
            // rect1 = top, rect2 = bottom
            if (addFillLeft) {
                rect1 = pathItems.rectangle(objTop, objLeft, objWidth, objTop - splitY);
                rect1.fillColor = colorLeft;
                rect1.stroked = false;
            }
            if (addFillRight) {
                rect2 = pathItems.rectangle(splitY, objLeft, objWidth, splitY - objBottom);
                rect2.fillColor = colorRight;
                rect2.stroked = false;
            }
            if (cornerRadius && Number(cornerRadius) > 0) {
                if (rect1) applyFillCornerRounding(rect1, "top", cornerRadius);
                if (rect2) applyFillCornerRounding(rect2, "bottom", cornerRadius);
            }
        } else {
            // rect1 = left, rect2 = right
            if (addFillLeft) {
                rect1 = pathItems.rectangle(objTop, objLeft, splitX - objLeft, objHeight);
                rect1.fillColor = colorLeft;
                rect1.stroked = false;
            }
            if (addFillRight) {
                rect2 = pathItems.rectangle(objTop, splitX, objRight - splitX, objHeight);
                rect2.fillColor = colorRight;
                rect2.stroked = false;
            }
            if (cornerRadius && Number(cornerRadius) > 0) {
                if (rect1) applyFillCornerRounding(rect1, "left", cornerRadius);
                if (rect2) applyFillCornerRounding(rect2, "right", cornerRadius);
            }
        }

        var frameRect = null;
        if (addOverallFrame) {
            try {
                frameRect = pathItems.rectangle(objTop, objLeft, objWidth, objHeight);
                frameRect.filled = false;
                frameRect.stroked = true;
                frameRect.strokeWidth = strokeWidthValue;
                frameRect.strokeColor = colorStroke;

                applyRoundCornersIfNeeded(frameRect, cornerRadius);

                try { frameRect.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (eFZ) { }
            } catch (eFrame) {
                frameRect = null;
            }
        }

        var dividerLine = null;
        if (addDivider) {
            try {
                dividerLine = pathItems.add();
                dividerLine.filled = false;
                dividerLine.stroked = true;
                dividerLine.strokeWidth = strokeWidthValue;
                dividerLine.strokeColor = colorStroke;

                if (isVertical) {
                    dividerLine.setEntirePath([
                        [objLeft, splitY],
                        [objRight, splitY]
                    ]);
                } else {
                    dividerLine.setEntirePath([
                        [splitX, objTop],
                        [splitX, objBottom]
                    ]);
                }

                try { dividerLine.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (eDZ) { }
            } catch (eDiv) {
                dividerLine = null;
            }
        }

        try { if (rect1) rect1.zOrder(ZOrderMethod.SENDTOBACK); } catch (eZ1) { }
        try { if (rect2) rect2.zOrder(ZOrderMethod.SENDTOBACK); } catch (eZ2) { }

        try { targetItem.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (eT1) { }

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
    function showHeightDialog(previewFn, clearPreviewFn) {
        var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
        dlg.orientation = 'column';
        dlg.alignChildren = ['fill', 'top'];
        dlg.margins = 18;

        try { PreviewHistory.start(); } catch (ePH0) { }

        var ss = getSessionSettings();

        setDialogOpacity(dlg, DIALOG_OPACITY);
        shiftDialogPosition(dlg, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);

        // --- プリセット / Presets ---
        var PRESETS_KEY = "SplitBackgroundForTwo_presets";
        if (!$.global[PRESETS_KEY]) $.global[PRESETS_KEY] = {};

        function getPresetNames() {
            var names = [];
            var store = $.global[PRESETS_KEY];
            for (var k in store) {
                if (store.hasOwnProperty(k)) names.push(k);
            }
            names.sort();
            return names;
        }

        function getCurrentPresetData() {
            var wo = Number(slWidth.value);
            if (isNaN(wo)) wo = 0;
            return {
                splitDirection: rbSplitLR.value ? 'lr' : 'tb',
                fillLeft: !!cbFillLeft.value,
                fillRight: !!cbFillRight.value,
                colorLeft: String(etColorLeft.text),
                colorRight: String(etColorRight.text),
                overallFrame: !!cbOverallFrame.value,
                divider: !!cbDivider.value,
                strokeUnit: String(etStroke.text),
                strokeColor: String(etStrokeColor.text),
                cornerEnabled: !!cbCorner.value,
                cornerUnit: String(etCorner.text),
                cornerAuto: !!cbCornerAuto.value,
                widthOffset: Math.round(wo * 10) / 10
            };
        }

        function applyPresetData(data) {
            if (!data) return;
            try {
                if (data.splitDirection === 'tb') { rbSplitTB.value = true; rbSplitLR.value = false; }
                else { rbSplitLR.value = true; rbSplitTB.value = false; }
                cbFillLeft.value = !!data.fillLeft;
                cbFillRight.value = !!data.fillRight;
                etColorLeft.text = data.colorLeft || 'DCDCDC';
                etColorRight.text = data.colorRight || '808080';
                cbOverallFrame.value = !!data.overallFrame;
                cbDivider.value = !!data.divider;
                etStroke.text = data.strokeUnit || '1';
                etStrokeColor.text = data.strokeColor || '000000';
                cbCorner.value = !!data.cornerEnabled;
                etCorner.text = data.cornerUnit || '0';
                cbCornerAuto.value = !!data.cornerAuto;
                updateBalanceLabels();
                updateStrokeWidthEnabled();
                updateCornerState();
                try { updateWidthSliderMaxBySelection(); } catch (e) { }
                if (data.widthOffset !== undefined) {
                    syncAllFromOffset(data.widthOffset);
                }
                applyPreview();
            } catch (e) { }
        }

        function refreshPresetDropdown() {
            ddPreset.removeAll();
            ddPreset.add('item', '---');
            var names = getPresetNames();
            for (var i = 0; i < names.length; i++) {
                ddPreset.add('item', names[i]);
            }
            ddPreset.selection = 0;
        }

        function exportCurrentSettingsToDesktop() {
            var d = getCurrentPresetData();
            var lines = [];
            lines.push('splitDirection=' + (d.splitDirection || 'lr'));
            lines.push('fillLeft=' + (d.fillLeft ? '1' : '0'));
            lines.push('fillRight=' + (d.fillRight ? '1' : '0'));
            lines.push('colorLeft=' + (d.colorLeft || 'DCDCDC'));
            lines.push('colorRight=' + (d.colorRight || '808080'));
            lines.push('overallFrame=' + (d.overallFrame ? '1' : '0'));
            lines.push('divider=' + (d.divider ? '1' : '0'));
            lines.push('strokeUnit=' + (d.strokeUnit || '1'));
            lines.push('strokeColor=' + (d.strokeColor || '000000'));
            lines.push('cornerEnabled=' + (d.cornerEnabled ? '1' : '0'));
            lines.push('cornerUnit=' + (d.cornerUnit || '0'));
            lines.push('cornerAuto=' + (d.cornerAuto ? '1' : '0'));
            lines.push('widthOffset=' + (d.widthOffset || 0));
            var desktopPath = Folder.desktop.fsName;
            var baseName = 'SplitBackgroundForTwo_settings';
            var ext = '.txt';
            var filePath = desktopPath + '/' + baseName + ext;
            var n = 1;
            while (new File(filePath).exists) {
                filePath = desktopPath + '/' + baseName + '_' + n + ext;
                n++;
            }
            var f = new File(filePath);
            f.encoding = 'UTF-8';
            if (f.open('w')) {
                f.write(lines.join('\n'));
                f.close();
                alert(L('alertExportedSettings') + '\n' + filePath);
            } else {
                alert(L('alertExportFailed'));
            }
        }

        function buildPresetRow(parent) {
            var row = parent.add('group');
            row.orientation = 'row';
            row.alignChildren = ['left', 'center'];
            row.alignment = ['fill', 'top'];

            row.add('statictext', undefined, L('panelPreset'));
            var dropdown = row.add('dropdownlist', undefined, ['---']);
            dropdown.preferredSize = [120, -1];
            dropdown.selection = 0;

            var btnSave = row.add('button', undefined, L('labelSave'));
            btnSave.preferredSize = [50, -1];

            var btnExport = row.add('button', undefined, L('labelExport'));
            btnExport.preferredSize = [60, -1];

            (function () {
                var names = getPresetNames();
                for (var i = 0; i < names.length; i++) {
                    dropdown.add('item', names[i]);
                }
            })();

            dropdown.onChange = function () {
                if (!dropdown.selection || dropdown.selection.index === 0) return;
                var name = dropdown.selection.text;
                var store = $.global[PRESETS_KEY];
                if (store[name]) applyPresetData(store[name]);
            };

            btnSave.onClick = function () {
                var nameDialog = new Window('dialog', L('presetSaveTitle'));
                nameDialog.orientation = 'column';
                nameDialog.alignChildren = ['fill', 'top'];
                nameDialog.add('statictext', undefined, L('presetNamePrompt'));
                var etName = nameDialog.add('edittext', undefined, '');
                etName.characters = 20;
                etName.active = true;
                var nameBtns = nameDialog.add('group');
                nameBtns.alignment = ['right', 'center'];
                nameBtns.add('button', undefined, L('labelCancel'), { name: 'cancel' });
                nameBtns.add('button', undefined, L('labelOK'), { name: 'ok' });
                if (nameDialog.show() !== 1) return;
                var name = etName.text;
                if (!name || name === '' || name === '---') return;
                $.global[PRESETS_KEY][name] = getCurrentPresetData();
                refreshPresetDropdown();
            };

            btnExport.onClick = function () {
                exportCurrentSettingsToDesktop();
            };

            return {
                row: row,
                ddPreset: dropdown,
                btnPresetSave: btnSave,
                btnPresetExport: btnExport
            };
        }

        function buildSplitDirectionPanel(parent) {
            var panel = parent.add('panel', undefined, L('panelSplitDirection'));
            panel.orientation = 'row';
            panel.alignChildren = ['left', 'center'];
            panel.margins = [15, 20, 15, 10];
            panel.alignment = ['fill', 'top'];

            var rbLR = panel.add('radiobutton', undefined, L('labelSplitLR'));
            var rbTB = panel.add('radiobutton', undefined, L('labelSplitTB'));
            var savedSplitDir = (ss.splitDirection !== undefined) ? ss.splitDirection : 'lr';
            rbLR.value = (savedSplitDir === 'lr');
            rbTB.value = (savedSplitDir === 'tb');
            if (!rbLR.value && !rbTB.value) rbLR.value = true;

            return {
                panel: panel,
                rbSplitLR: rbLR,
                rbSplitTB: rbTB
            };
        }

        function buildFillPanel(parent) {
            var panel = parent.add('panel', undefined, L('panelFill'));
            panel.orientation = 'column';
            panel.alignChildren = ['left', 'top'];
            panel.margins = [15, 20, 15, 10];
            panel.alignment = ['fill', 'top'];

            var fillLeftRow = panel.add('group');
            fillLeftRow.orientation = 'row';
            fillLeftRow.alignChildren = ['left', 'center'];
            var cbLeft = fillLeftRow.add('checkbox', undefined, L('labelFillLeft'));
            cbLeft.value = (ss.fillLeft !== undefined) ? !!ss.fillLeft : true;
            var etLeft = fillLeftRow.add('edittext', undefined, (ss.colorLeft !== undefined) ? ss.colorLeft : 'DCDCDC');
            etLeft.characters = 16;
            createColorSwatch(fillLeftRow, etLeft, applyPreview);

            var fillRightRow = panel.add('group');
            fillRightRow.orientation = 'row';
            fillRightRow.alignChildren = ['left', 'center'];
            var cbRight = fillRightRow.add('checkbox', undefined, L('labelFillRight'));
            cbRight.value = (ss.fillRight !== undefined) ? !!ss.fillRight : true;
            var etRight = fillRightRow.add('edittext', undefined, (ss.colorRight !== undefined) ? ss.colorRight : '808080');
            etRight.characters = 16;
            createColorSwatch(fillRightRow, etRight, applyPreview);

            return {
                panel: panel,
                cbFillLeft: cbLeft,
                etColorLeft: etLeft,
                cbFillRight: cbRight,
                etColorRight: etRight
            };
        }

        function buildLinePanel(parent) {
            var panel = parent.add('panel', undefined, L('panelLine'));
            panel.orientation = 'column';
            panel.alignChildren = ['left', 'top'];
            panel.margins = [15, 20, 15, 10];
            panel.alignment = ['fill', 'top'];

            var lineCbRow = panel.add('group');
            lineCbRow.orientation = 'row';
            lineCbRow.alignChildren = ['left', 'center'];

            var cbFrame = lineCbRow.add('checkbox', undefined, L('labelOverallFrame'));
            cbFrame.value = (ss.overallFrame !== undefined) ? !!ss.overallFrame : false;

            var cbDiv = lineCbRow.add('checkbox', undefined, L('labelDivider'));
            cbDiv.value = (ss.divider !== undefined) ? !!ss.divider : false;

            var strokeRow = panel.add('group');
            strokeRow.orientation = 'row';
            strokeRow.alignChildren = ['left', 'center'];

            strokeRow.add('statictext', undefined, L('labelStrokeWidth'));
            var etStrokeLocal = strokeRow.add('edittext', undefined, (ss && ss.strokeUnit !== undefined) ? String(ss.strokeUnit) : formatUnitValue(ptToUnit(1, "strokeUnits")));
            etStrokeLocal.characters = 3;
            changeValueByArrowKey(etStrokeLocal, false, applyPreview);
            strokeRow.add('statictext', undefined, getCurrentUnitLabelByPrefKey("strokeUnits"));

            var colorRow = panel.add('group');
            colorRow.orientation = 'row';
            colorRow.alignChildren = ['left', 'center'];
            colorRow.add('statictext', undefined, L('labelColor'));
            var etStrokeColorLocal = colorRow.add('edittext', undefined, (ss.strokeColor !== undefined) ? ss.strokeColor : '000000');
            etStrokeColorLocal.characters = 16;
            createColorSwatch(colorRow, etStrokeColorLocal, applyPreview);

            return {
                panel: panel,
                cbOverallFrame: cbFrame,
                cbDivider: cbDiv,
                lineRow: strokeRow,
                etStroke: etStrokeLocal,
                lineColorRow: colorRow,
                etStrokeColor: etStrokeColorLocal
            };
        }

        function buildOptionPanel(parent) {
            var panel = parent.add('panel', undefined, L('panelOptions'));
            panel.orientation = 'column';
            panel.alignChildren = ['left', 'top'];
            panel.margins = [15, 20, 15, 10];
            panel.alignment = ['fill', 'top'];

            var row = panel.add('group');
            row.orientation = 'row';
            row.alignChildren = ['left', 'center'];

            var cbCornerLocal = row.add('checkbox', undefined, L('labelCornerRadius'));
            cbCornerLocal.value = (ss.cornerEnabled !== undefined) ? !!ss.cornerEnabled : false;
            var defaultCornerUnit = (ss && ss.cornerUnit !== undefined) ? String(ss.cornerUnit) : formatUnitValue(ptToUnit(0, "rulerType"));
            var etCornerLocal = row.add('edittext', undefined, defaultCornerUnit);
            etCornerLocal.characters = 4;
            row.add('statictext', undefined, getCurrentUnitLabelByPrefKey("rulerType"));
            changeValueByArrowKey(etCornerLocal, false, applyPreview);

            var cbCornerAutoLocal = row.add('checkbox', undefined, L('labelPillShape'));
            cbCornerAutoLocal.value = (ss.cornerAuto !== undefined) ? !!ss.cornerAuto : false;

            return {
                panel: panel,
                cornerRow: row,
                cbCorner: cbCornerLocal,
                etCorner: etCornerLocal,
                cbCornerAuto: cbCornerAutoLocal
            };
        }

        function buildBalancePanel(parent) {
            var panel = parent.add('panel', undefined, L('panelFixed'));
            panel.orientation = 'column';
            panel.alignChildren = ['fill', 'top'];
            panel.margins = [15, 20, 15, 10];
            panel.alignment = ['fill', 'top'];

            var pinWidthColLocal = panel.add('group');
            pinWidthColLocal.orientation = 'column';
            pinWidthColLocal.alignChildren = ['fill', 'top'];
            pinWidthColLocal.spacing = 6;

            var unitLabelLocal = getCurrentUnitLabelByPrefKey("rulerType");

            var leftRowLocal = pinWidthColLocal.add('group');
            leftRowLocal.orientation = 'row';
            leftRowLocal.alignChildren = ['left', 'center'];
            var stLeftLabelLocal = leftRowLocal.add('statictext', undefined, L('labelFillLeft'));
            var etLeftWidthLocal = leftRowLocal.add('edittext', undefined, '0');
            etLeftWidthLocal.characters = 4;
            leftRowLocal.add('statictext', undefined, unitLabelLocal + '、');
            var etLeftPctLocal = leftRowLocal.add('edittext', undefined, '50');
            etLeftPctLocal.characters = 3;
            leftRowLocal.add('statictext', undefined, '%');

            var rightRowLocal = pinWidthColLocal.add('group');
            rightRowLocal.orientation = 'row';
            rightRowLocal.alignChildren = ['left', 'center'];
            var stRightLabelLocal = rightRowLocal.add('statictext', undefined, L('labelFillRight'));
            var etRightWidthLocal = rightRowLocal.add('edittext', undefined, '0');
            etRightWidthLocal.characters = 4;
            rightRowLocal.add('statictext', undefined, unitLabelLocal + '、');
            var etRightPctLocal = rightRowLocal.add('edittext', undefined, '50');
            etRightPctLocal.characters = 3;
            rightRowLocal.add('statictext', undefined, '%');

            var sliderRowLocal = pinWidthColLocal.add('group');
            sliderRowLocal.orientation = 'row';
            sliderRowLocal.alignChildren = ['fill', 'center'];

            var slWidthLocal = sliderRowLocal.add('slider', undefined, 0, 0, 0);
            slWidthLocal.preferredSize = [180, 20];

            return {
                panel: panel,
                pinWidthCol: pinWidthColLocal,
                leftRow: leftRowLocal,
                stLeftLabel: stLeftLabelLocal,
                etLeftWidth: etLeftWidthLocal,
                etLeftPct: etLeftPctLocal,
                rightRow: rightRowLocal,
                stRightLabel: stRightLabelLocal,
                etRightWidth: etRightWidthLocal,
                etRightPct: etRightPctLocal,
                pinWidthSliderRow: sliderRowLocal,
                slWidth: slWidthLocal,
                unitLabel: unitLabelLocal
            };
        }

        var presetUI = buildPresetRow(dlg);
        var presetRow = presetUI.row;
        var ddPreset = presetUI.ddPreset;
        var btnPresetSave = presetUI.btnPresetSave;
        var btnPresetExport = presetUI.btnPresetExport;

        // --- 分割方向 / Split direction ---
        var splitDirUI = buildSplitDirectionPanel(dlg);
        var splitDirPanel = splitDirUI.panel;
        var rbSplitLR = splitDirUI.rbSplitLR;
        var rbSplitTB = splitDirUI.rbSplitTB;

        // --- 描画パネル / Drawing panel ---
        var fillUI = buildFillPanel(dlg);
        var drawPanel = fillUI.panel;
        var cbFillLeft = fillUI.cbFillLeft;
        var etColorLeft = fillUI.etColorLeft;
        var cbFillRight = fillUI.cbFillRight;
        var etColorRight = fillUI.etColorRight;

        // --- 線パネル / Stroke panel ---
        var lineUI = buildLinePanel(dlg);
        var linePanel = lineUI.panel;
        var cbOverallFrame = lineUI.cbOverallFrame;
        var cbDivider = lineUI.cbDivider;
        var lineRow = lineUI.lineRow;
        var etStroke = lineUI.etStroke;
        var lineColorRow = lineUI.lineColorRow;
        var etStrokeColor = lineUI.etStrokeColor;

        function updateStrokeWidthEnabled() {
            var on = !!(cbOverallFrame.value || cbDivider.value);
            try { lineRow.enabled = on; } catch (eEn1) { }
            try { lineColorRow.enabled = on; } catch (eEn3) { }
        }

        updateStrokeWidthEnabled();

        // --- オプションパネル / Options panel ---
        var optionUI = buildOptionPanel(dlg);
        var optPanel = optionUI.panel;
        var cornerRow = optionUI.cornerRow;
        var cbCorner = optionUI.cbCorner;
        var etCorner = optionUI.etCorner;
        var cbCornerAuto = optionUI.cbCornerAuto;

        function updateCornerState() {
            try { etCorner.enabled = cbCorner.value && !cbCornerAuto.value; } catch (eEn2) { }
            try { cbCornerAuto.enabled = cbCorner.value; } catch (eEn3) { }
            if (cbCorner.value && cbCornerAuto.value && maxCornerUnit > 0) {
                etCorner.text = String(maxCornerUnit);
            }
        }

        cbCorner.onClick = function () { updateCornerState(); applyPreview(); };
        cbCornerAuto.onClick = function () { updateCornerState(); applyPreview(); };
        updateCornerState();

        // --- バランス / Balance (full-width, spanning both columns) ---
        var balanceUI = buildBalancePanel(dlg);
        var pinPanel = balanceUI.panel;
        var pinWidthCol = balanceUI.pinWidthCol;
        var leftRow = balanceUI.leftRow;
        var stLeftLabel = balanceUI.stLeftLabel;
        var etLeftWidth = balanceUI.etLeftWidth;
        var etLeftPct = balanceUI.etLeftPct;
        var rightRow = balanceUI.rightRow;
        var stRightLabel = balanceUI.stRightLabel;
        var etRightWidth = balanceUI.etRightWidth;
        var etRightPct = balanceUI.etRightPct;
        var pinWidthSliderRow = balanceUI.pinWidthSliderRow;
        var slWidth = balanceUI.slWidth;

        function getMaxV() {
            var maxV = 0;
            try { maxV = Number(slWidth.maxvalue); } catch (e) { }
            if (isNaN(maxV) || maxV <= 0) maxV = 0;
            return maxV;
        }

        function clampOffset(v) {
            v = Number(v);
            if (isNaN(v)) return 0;
            var maxV = getMaxV();
            if (v > maxV) v = maxV;
            if (v < -maxV) v = -maxV;
            return Math.round(v * 10) / 10;
        }

        // スライダー値(offset)から全フィールドを更新
        function syncAllFromOffset(v) {
            v = clampOffset(v);
            slWidth.value = v;
            var maxV = getMaxV();
            var lw = Math.round((maxV + v) * 10) / 10;
            var rw = Math.round((maxV - v) * 10) / 10;
            etLeftWidth.text = String(lw);
            etRightWidth.text = String(rw);
            if (maxV > 0) {
                etLeftPct.text = String(Math.round((maxV + v) / (2 * maxV) * 100));
                etRightPct.text = String(Math.round((maxV - v) / (2 * maxV) * 100));
            } else {
                etLeftPct.text = '50';
                etRightPct.text = '50';
            }
        }

        function syncFromSlider() {
            var kb = ScriptUI.environment.keyboardState;
            var v = Number(slWidth.value);
            if (kb && kb.altKey) v = Math.round(v);
            syncAllFromOffset(v);
        }

        function syncFromLeftWidth() {
            var lw = Number(etLeftWidth.text);
            if (isNaN(lw)) lw = 0;
            syncAllFromOffset(lw - getMaxV());
        }

        function syncFromRightWidth() {
            var rw = Number(etRightWidth.text);
            if (isNaN(rw)) rw = 0;
            syncAllFromOffset(getMaxV() - rw);
        }

        function syncFromLeftPct() {
            var pct = Number(etLeftPct.text);
            if (isNaN(pct)) pct = 50;
            if (pct < 0) pct = 0;
            if (pct > 100) pct = 100;
            var maxV = getMaxV();
            syncAllFromOffset((pct / 100 - 0.5) * 2 * maxV);
        }

        function syncFromRightPct() {
            var pct = Number(etRightPct.text);
            if (isNaN(pct)) pct = 50;
            if (pct < 0) pct = 0;
            if (pct > 100) pct = 100;
            var maxV = getMaxV();
            syncAllFromOffset((0.5 - pct / 100) * 2 * maxV);
        }

        changeValueByArrowKey(etLeftWidth, false, function () { syncFromLeftWidth(); applyPreview(); });
        changeValueByArrowKey(etRightWidth, false, function () { syncFromRightWidth(); applyPreview(); });
        changeValueByArrowKey(etLeftPct, false, function () { syncFromLeftPct(); applyPreview(); });
        changeValueByArrowKey(etRightPct, false, function () { syncFromRightPct(); applyPreview(); });

        etLeftWidth.addEventListener('changing', function () { syncFromLeftWidth(); applyPreview(); });
        etRightWidth.addEventListener('changing', function () { syncFromRightWidth(); applyPreview(); });
        etLeftPct.addEventListener('changing', function () { syncFromLeftPct(); applyPreview(); });
        etRightPct.addEventListener('changing', function () { syncFromRightPct(); applyPreview(); });

        var maxCornerUnit = 0; // オブジェクト高さの1/2（角丸上限）

        function updateWidthSliderMaxBySelection() {
            var maxUnit = 0;
            var tempGroupForMax = null;

            try {
                var hostLayer2 = null;
                try { hostLayer2 = item.layer; } catch (eHLm) { hostLayer2 = null; }
                if (!hostLayer2) {
                    try { hostLayer2 = doc.activeLayer; } catch (eHLm2) { hostLayer2 = null; }
                }
                if (!hostLayer2) hostLayer2 = doc.layers[0];

                unlockAndShow(hostLayer2);
                tempGroupForMax = hostLayer2.groupItems.add();
                tempGroupForMax.name = TEMP_NAME_PREFIX + "__temp_outline_group_for_max__";
                markTemp(tempGroupForMax);
                try { tempGroupForMax.opacity = 0; } catch (eOpM) { }

                var bi = getBoundsInfo(item, tempGroupForMax);
                var b = bi.bounds;

                var widthPt = b[2] - b[0];
                if (widthPt < 0) widthPt = 0;
                var heightPt = b[1] - b[3];
                if (heightPt < 0) heightPt = 0;

                // スライダー範囲: 分割方向に応じて幅or高さの半分
                var dimPt = rbSplitTB.value ? heightPt : widthPt;
                maxUnit = ptToUnit(dimPt / 2, "rulerType");
                if (isNaN(maxUnit) || maxUnit < 0) maxUnit = 0;
                maxUnit = Math.round(maxUnit * 10) / 10;

                // 角丸上限: 分割方向に応じて短辺の1/2
                var cornerDimPt = rbSplitTB.value ? widthPt : heightPt;
                maxCornerUnit = ptToUnit(cornerDimPt / 2, "rulerType");
                if (isNaN(maxCornerUnit) || maxCornerUnit < 0) maxCornerUnit = 0;
                maxCornerUnit = Math.round(maxCornerUnit * 10) / 10;

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

            try { slWidth.minvalue = -maxUnit; } catch (eSetMin) { }
            try { slWidth.maxvalue = maxUnit; } catch (eSetMax) { }

            syncAllFromOffset(0);
        }

        slWidth.onChanging = function () {
            syncFromSlider();
            applyPreview();
        };

        // プレビューは常にON
        var cbPreview = { value: true };

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
            if (!cbCorner.value) return 0;
            if (cbCornerAuto.value && maxCornerUnit > 0) {
                etCorner.text = String(maxCornerUnit);
            }
            var vUnit = Number(etCorner.text);
            if (isNaN(vUnit) || vUnit < 0) return null;
            vUnit = Math.round(vUnit * 10) / 10;
            if (maxCornerUnit > 0 && vUnit > maxCornerUnit) {
                vUnit = maxCornerUnit;
                etCorner.text = String(vUnit);
            }
            var vPt = unitToPt(vUnit, "rulerType");
            if (isNaN(vPt) || vPt < 0) return null;
            vPt = Math.round(vPt * 10) / 10;
            return vPt;
        }

        function parseWidthOffset() {
            var vUnit = Number(slWidth.value);
            if (isNaN(vUnit)) vUnit = 0;
            vUnit = Math.round(vUnit * 10) / 10;

            var vPt = unitToPt(vUnit, "rulerType");
            if (isNaN(vPt)) vPt = 0;
            vPt = Math.round(vPt * 10) / 10;
            return vPt;
        }

        function applyPreview() {
            // カラーフィールドから色を更新（RGB/CMYK対応）
            colorLeft = parseColorString(etColorLeft.text);
            colorRight = parseColorString(etColorRight.text);
            colorStroke = parseColorString(etStrokeColor.text);

            var sw = parseStrokeWidth();
            if (sw === null) {
                try { removeMarkedTempItems(doc); } catch (eTmp0) { }
                clearPreviewItemsOnly(doc);
                clearPreviewFn();
                return;
            }

            var cr = parseCornerRadius();
            if (cr === null) {
                try { removeMarkedTempItems(doc); } catch (eTmp0) { }
                clearPreviewItemsOnly(doc);
                clearPreviewFn();
                return;
            }

            var wo = parseWidthOffset();

            var sd = rbSplitLR.value ? 'lr' : 'tb';

            previewFn(
                cbPreview.value,
                cbOverallFrame.value,
                cbDivider.value,
                cbFillLeft.value,
                cbFillRight.value,
                sw,
                cr,
                wo,
                sd
            );
        }

        function persistCurrentUIToSession() {
            try {
                var s = getSessionSettings();
                s.preview = !!cbPreview.value;
                s.overallFrame = !!cbOverallFrame.value;
                s.divider = !!cbDivider.value;
                s.fillLeft = !!cbFillLeft.value;
                s.fillRight = !!cbFillRight.value;

                s.strokeUnit = String(etStroke.text);
                s.cornerUnit = String(etCorner.text);
                s.cornerEnabled = !!cbCorner.value;
                s.cornerAuto = !!cbCornerAuto.value;

                s.widthUnit = Number(slWidth.value);
                if (isNaN(s.widthUnit)) s.widthUnit = 0;
                s.widthUnit = Math.round(s.widthUnit * 10) / 10;

                s.splitDirection = rbSplitLR.value ? 'lr' : 'tb';
                s.colorLeft = String(etColorLeft.text);
                s.colorRight = String(etColorRight.text);
                s.strokeColor = String(etStrokeColor.text);

                saveSessionSettings(s);
            } catch (e) { }
        }

        (function () {
            var prevOnShow = dlg.onShow;
            dlg.onShow = function () {
                if (typeof prevOnShow === "function") {
                    try { prevOnShow(); } catch (ePrevShow) { }
                }
                // 元のオブジェクトを非表示
                try { item.hidden = true; } catch (eHide) { }
                // 幅と高さから分割方向を自動判定
                try {
                    var gb = item.geometricBounds; // [left, top, right, bottom]
                    var w = gb[2] - gb[0];
                    var h = gb[1] - gb[3];
                    if (h > w) {
                        rbSplitTB.value = true;
                        rbSplitLR.value = false;
                    } else {
                        rbSplitLR.value = true;
                        rbSplitTB.value = false;
                    }
                } catch (eAutoDir) { }
                try { updateBalanceLabels(); } catch (eLbl) { }
                try { updateWidthSliderMaxBySelection(); } catch (eMaxInit) { }
                // スライダーを中央（0）に設定
                try { slWidth.value = 0; syncAllFromOffset(0); } catch (eSlInit) { }
                try { applyPreview(); } catch (eInitPrev) { }
            };
        })();

        etStroke.addEventListener('changing', function () { applyPreview(); });
        etCorner.addEventListener('changing', function () { applyPreview(); });
        etColorLeft.addEventListener('changing', function () { applyPreview(); });
        etColorRight.addEventListener('changing', function () { applyPreview(); });
        etStrokeColor.addEventListener('changing', function () { applyPreview(); });

        function updateBalanceLabels() {
            var isVert = rbSplitTB.value;
            try { stLeftLabel.text = isVert ? '上' : '左'; } catch (e) { }
            try { stRightLabel.text = isVert ? '下' : '右'; } catch (e) { }
        }

        rbSplitLR.onClick = function () { updateBalanceLabels(); try { updateWidthSliderMaxBySelection(); } catch (e) { } applyPreview(); };
        rbSplitTB.onClick = function () { updateBalanceLabels(); try { updateWidthSliderMaxBySelection(); } catch (e) { } applyPreview(); };
        cbOverallFrame.onClick = function () { updateStrokeWidthEnabled(); applyPreview(); };
        cbDivider.onClick = function () { updateStrokeWidthEnabled(); applyPreview(); };
        cbFillLeft.onClick = function () { applyPreview(); };
        cbFillRight.onClick = function () { applyPreview(); };

        var btns = dlg.add('group');
        btns.orientation = 'row';
        btns.alignment = ['right', 'center'];

        var cancelBtn = btns.add('button', undefined, L('labelCancel'), { name: 'cancel' });
        var okBtn = btns.add('button', undefined, L('labelOK'), { name: 'ok' });

        okBtn.onClick = function () {
            try { removeMarkedTempItems(doc); } catch (eTmpOK) { }
            clearPreviewItemsOnly(doc);
            try { removeMarkedPreviewItems(doc); } catch (ePrevRmOK) { }
            try { removePreviewLayerIfExists(doc); } catch (eRmPL_OK) { }

            previewState.rect1 = result.rect1;
            previewState.rect2 = result.rect2;
            previewState.frame = result.frame;
            previewState.divider = result.divider;
            previewState.applied = false;

            persistCurrentUIToSession();
            dlg.close(1);
        };

        cancelBtn.onClick = function () {
            try { removeMarkedTempItems(doc); } catch (eTmpCancel) { }
            clearPreviewItemsOnly(doc);
            clearPreviewFn();
            try { removePreviewLayerIfExists(doc); } catch (eRmPL_Cancel) { }
            // 元のオブジェクトを再表示
            try { item.hidden = false; } catch (eShow) { }
            persistCurrentUIToSession();
            dlg.close(0);
        };

        dlg.defaultElement = okBtn;
        dlg.cancelElement = cancelBtn;

        dlg.onClose = function () {
            try { removeMarkedTempItems(doc); } catch (eTmpClose) { }
            clearPreviewItemsOnly(doc);
            try { removePreviewLayerIfExists(doc); } catch (eRmPL_Close) { }
            // 元のオブジェクトを再表示（キャンセル・閉じる時の安全策）
            try { item.hidden = false; } catch (eShow) { }
            persistCurrentUIToSession();
            return true;
        };

        var result = dlg.show();
        if (result !== 1) return null;

        return true;
    }

    function previewFn(enabled, addOverallFrame, addDivider, addFillLeft, addFillRight, strokeWidth, cornerRadius, widthOffset, splitDir) {
        // Always remove previous preview by deleting preview-marked items (do NOT rely on Undo).
        try { removeMarkedTempItems(doc); } catch (eTmp0) { }
        clearPreviewItemsOnly(doc);

        // Create/ensure preview layer and clear it every refresh
        var previewLayer = null;
        try {
            var la = null;
            try { la = item.layer; } catch (eLA) { la = null; }
            previewLayer = ensurePreviewLayer(doc, la, la);
            clearLayerItems(previewLayer);
        } catch (ePL) { previewLayer = null; }

        if (!enabled) return;

        try {
            var result = drawBackgroundRects(item, addOverallFrame, addDivider, addFillLeft, addFillRight, strokeWidth, cornerRadius, widthOffset, splitDir);

            previewState.rect1 = result.rect1;
            previewState.rect2 = result.rect2;
            previewState.frame = result.frame;
            previewState.divider = result.divider;
            previewState.applied = true;

            if (previewState.rect1) markPreview(previewState.rect1);
            if (previewState.rect2) markPreview(previewState.rect2);
            if (previewState.frame) markPreview(previewState.frame);
            if (previewState.divider) markPreview(previewState.divider);

            previewApplied = true;
            previewState.last.overallFrame = !!addOverallFrame;
            previewState.last.divider = !!addDivider;
            previewState.last.fillLeft = !!addFillLeft;
            previewState.last.fillRight = !!addFillRight;
            previewState.last.strokeWidth = strokeWidth;
            previewState.last.cornerRadius = cornerRadius;
            previewState.last.widthOffset = (widthOffset !== undefined && widthOffset !== null) ? widthOffset : 0;
            previewState.last.splitDirection = splitDir || 'lr';

            safeRedraw();
        } catch (ePrev) {
            try { clearPreviewRects(); } catch (eClr) { }
            throw ePrev;
        }
    }

    var dialogResult = showHeightDialog(previewFn, clearPreviewRects);
    if (dialogResult === null) return;

    // --- 確定処理 ---
    // ［OK］時は PreviewHistory.undo() によりプレビュー生成物は消えているため、
    // 必ずここで最終描画を行い「OK直前の見た目」で確定する。
    var finalResult = drawBackgroundRects(
        item,
        previewState.last.overallFrame, previewState.last.divider,
        previewState.last.fillLeft, previewState.last.fillRight,
        previewState.last.strokeWidth, previewState.last.cornerRadius,
        previewState.last.widthOffset, previewState.last.splitDirection
    );
    previewState.rect1 = finalResult.rect1;
    previewState.rect2 = finalResult.rect2;
    previewState.frame = finalResult.frame;
    previewState.divider = finalResult.divider;

    // 確定描画はプレビュー扱いにしない
    if (previewState.rect1) unmarkPreview(previewState.rect1);
    if (previewState.rect2) unmarkPreview(previewState.rect2);
    if (previewState.frame) unmarkPreview(previewState.frame);
    if (previewState.divider) unmarkPreview(previewState.divider);

    // 元のオブジェクトを削除
    try { item.remove(); } catch (eRmItem) { }

    doc.selection = null;

    // Remove preview layer if it exists (keep document clean)
    try { removePreviewLayerIfExists(doc); } catch (eRmPL_Final) { }

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