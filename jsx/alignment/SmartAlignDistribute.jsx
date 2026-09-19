#target illustrator
#targetengine "SmartAlignAndTileEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトを縦または横に並べ、指定した間隔で分布します。
方向は自動判定でき、揃え（左右／上下）、プレビュー境界、ランダム並べ替えにも対応します。

詳細は README を参照してください。

### Overview

Lines the selected objects up vertically or horizontally and distributes them at the spacing you specify.
The direction can be detected automatically, and alignment, preview bounds and random reordering are all supported.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartAlignDistribute";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAlignDistribute.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignDistribute.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var DIALOG_OPACITY = 0.97;
    var DIALOG_POSITION_KEY = "__SmartAlignDistribute_DialogPosition__";
    var PREVIEW_MIN_INTERVAL_MS = 80;
    /* 計測用に一時的に作るグループの名前 / name of the throwaway measuring group */
    var TEMP_MEASURE_GROUP_NAME = "__SmartAlignDistribute_TempMeasure__";
    var PREVIEW_SCHEDULE_MS = 60;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 現在のロケールから言語を判定 / Detect language from locale */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    var LABELS = {
        /* === ダイアログ / Dialog === */
        dialog: {
            title: { ja: "整列と分布", en: "Align & Distribute" }
        },

        /* === 共通 / Common === */
        common: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },

        /* === 方向 / Direction === */
        direction: {
            panel: { ja: "方向", en: "Direction" },
            auto: { ja: "自動", en: "Auto" },
            vertical: { ja: "縦", en: "Vertical" },
            horizontal: { ja: "横", en: "Horizontal" },
            autoTip: {
                ja: "選択範囲が横長なら横並び、縦長なら縦並びとして扱います。",
                en: "Lays the objects out in a row when the selection is wider than tall, in a column otherwise."
            },
            verticalTip: { ja: "上から下へ縦に並べます。", en: "Stacks the objects from top to bottom." },
            horizontalTip: { ja: "左から右へ横に並べます。", en: "Lays the objects out from left to right." }
        },

        /* === 間隔 / Spacing === */
        spacing: {
            label: { ja: "間隔", en: "Spacing" },
            tip: {
                ja: "オブジェクト間のすき間。マイナス値で重ねられます。",
                en: "Gap between objects. Negative values overlap them."
            }
        },

        /* === 揃え / Alignment === */
        align: {
            horizontalTitle: { ja: "揃え（左右）", en: "Align (H)" },
            verticalTitle: { ja: "揃え（上下）", en: "Align (V)" },
            hLeft: { ja: "左", en: "Left" },
            hCenter: { ja: "中央", en: "Center" },
            hRight: { ja: "右", en: "Right" },
            hNone: { ja: "なし", en: "None" },
            vTop: { ja: "上", en: "Top" },
            vMiddle: { ja: "中央", en: "Middle" },
            vBottom: { ja: "下", en: "Bottom" },
            vNone: { ja: "なし", en: "None" },
            horizontalTip: {
                ja: "縦に並べたときの左右の揃え方です。N／L／C／R キーでも切り替えられます。",
                en: "Horizontal alignment used when stacking vertically. The keys N / L / C / R switch it."
            },
            verticalTip: {
                ja: "横に並べたときの上下の揃え方です。N／T／M／B キーでも切り替えられます。",
                en: "Vertical alignment used when laying out horizontally. The keys N / T / M / B switch it."
            }
        },

        /* === オプション / Options === */
        options: {
            useBounds: { ja: "プレビュー境界を使用", en: "Use preview bounds" },
            useBoundsTip: {
                ja: "線や効果を含む見た目の境界で整列します。オフはパスのみのジオメトリ境界。",
                en: "Align by visible bounds (incl. strokes/effects). Off uses geometric (path-only) bounds."
            },
            random: { ja: "ランダム", en: "Random" },
            randomTip: {
                ja: "並び順をランダムに入れ替えます（左上の位置は維持）。",
                en: "Shuffle the stacking order at random (top-left position is kept)."
            },
            measureText: { ja: "テキストの高さを計測", en: "Measure text height" },
            measureTextTip: {
                ja: "縦並び時のみ、テキストを一度だけ複製→アウトライン化して境界を計測します（ダイアログ中だけキャッシュ）。",
                en: "Only in vertical layout, measures text by duplicating and outlining once (cached for this dialog only)."
            }
        },

        /* === エラー / Errors === */
        error: {
            needSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
            prefix: { ja: "エラーが発生しました: ", en: "An error has occurred: " }
        }
    };

    /* ドットパスのキーで LABELS から現在言語の文字列を取得（{slash}→"/"） / Resolve dotted key in LABELS for current language ({slash}→"/") */
    function getLabel(labelPath) {
        var labelNode = LABELS;
        var pathKeys = labelPath.split(".");
        for (var i = 0; i < pathKeys.length; i++) {
            if (labelNode == null) return labelPath;
            labelNode = labelNode[pathKeys[i]];
        }
        if (labelNode == null) return labelPath;
        var text = (labelNode[uiLang] != null) ? labelNode[uiLang] : labelNode.en;
        if (text == null) return labelPath;
        return String(text).replace(/\{slash\}/g, "/");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /* プレビューを再生成（前回分を undo してから processFn 実行） / Re-render preview */
    function runPreview(state, processFn, isEnabled) {
        try {
            if (isEnabled) {
                if (state.isUndo) app.undo();
                else state.isUndo = true;
                processFn();
                app.redraw();
            } else if (state.isUndo) {
                app.undo();
                app.redraw();
                state.isUndo = false;
            }
        } catch (err) { }
    }

    /* 確定処理の直前にプレビュー分を巻き戻す / Undo preview before final commit */
    function undoPreview(state) {
        try {
            if (state.isUndo) app.undo();
        } catch (err) { }
        state.isUndo = false;
    }

    /* ダイアログクローズ時のクリーンアップ（プレビュー巻き戻し＋一時レイヤー削除） / Cleanup on dialog close (undo preview + remove temp layer) */
    function cleanupPreview(state, doc, tempLayerName) {
        try {
            if (state.isUndo) app.undo();
            state.isUndo = false;
        } catch (err) { }
        if (tempLayerName) {
            try {
                var tempLayer = doc.layers.getByName(tempLayerName);
                tempLayer.remove();
            } catch (err) { }
        }
    }

    /* ダイアログ位置をセッション内に復元 / Load dialog position within session */
    function loadDialogPosition() {
        var savedPosition = $.global[DIALOG_POSITION_KEY];
        return (savedPosition && savedPosition.length === 2) ? [savedPosition[0], savedPosition[1]] : null;
    }

    /* ダイアログ位置をセッション内に保存 / Save dialog position within session */
    function saveDialogPosition(dialogLocation) {
        if (!dialogLocation || dialogLocation.length !== 2) return;
        $.global[DIALOG_POSITION_KEY] = [Math.round(dialogLocation[0]), Math.round(dialogLocation[1])];
    }

    /* クリップグループのクリッピングパスを取得（無ければ null） / Get clipping path of a clip group (null if none) */
    function getClippingPath(targetItem) {
        if (!targetItem || targetItem.typename !== "GroupItem" || !targetItem.clipped) return null;

        var children = targetItem.pageItems;
        for (var i = 0; i < children.length; i++) {
            var child = children[i];
            /* clipping を持たない種類のアイテムが混ざると例外になる / some item kinds do not expose clipping */
            try {
                if (child.clipping === true) return child;
            } catch (e) { }
            /* 複合パスはマスク本体に clipping が無く、内部パスに付く / Compound path: clipping flag sits on inner path, not the wrapper */
            if (child.typename === "CompoundPathItem" && child.pathItems && child.pathItems.length) {
                try {
                    if (child.pathItems[0].clipping === true) return child;
                } catch (e) { }
            }
        }
        return null;
    }

    /* アイテムの境界を取得（クリップグループはクリッピングパスを対象） / Get item bounds (clip group → its clipping path) */
    function getItemBounds(pageItem, usePreviewBounds) {
        var clipPath = getClippingPath(pageItem);
        var measuredItem = clipPath ? clipPath : pageItem;
        return usePreviewBounds ? measuredItem.visibleBounds : measuredItem.geometricBounds;
    }

    /* 選択範囲の幅・高さスパンを取得 / Get the span of a selection */
    function getSelectionSpan(items, usePreviewBounds) {
        var minLeft = null, maxRight = null, maxTop = null, minBottom = null;
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) continue;
            var bounds = getItemBounds(item, usePreviewBounds);
            if (minLeft === null || bounds[0] < minLeft) minLeft = bounds[0];
            if (maxRight === null || bounds[2] > maxRight) maxRight = bounds[2];
            if (maxTop === null || bounds[1] > maxTop) maxTop = bounds[1];
            if (minBottom === null || bounds[3] < minBottom) minBottom = bounds[3];
        }
        if (minLeft === null) return { spanX: 0, spanY: 0 };
        return { spanX: (maxRight - minLeft), spanY: (maxTop - minBottom) };
    }

    /* 自動方向判定（geometricBoundsで安定判定） / Auto-detect direction */
    function detectDirection(items) {
        var span = getSelectionSpan(items, false);
        return (span.spanX >= span.spanY) ? "horizontal" : "vertical";
    }

    /* Y座標でソート（上→下、同じなら左→右） / Sort by Y (top to bottom, then left) */
    function sortByY(items) {
        var sorted = items.slice();
        sorted.sort(function (a, b) {
            if (a.top !== b.top) return b.top - a.top;
            return a.left - b.left;
        });
        return sorted;
    }

    /* X座標でソート（左→右、同じなら上→下） / Sort by X (left to right, then top) */
    function sortByX(items) {
        var sorted = items.slice();
        sorted.sort(function (a, b) {
            if (a.left !== b.left) return a.left - b.left;
            return b.top - a.top;
        });
        return sorted;
    }

    /* Fisher-Yatesでインデックス配列をシャッフル / Fisher-Yates shuffle of index array */
    function makeShuffledIndices(count) {
        var indices = [];
        for (var i = 0; i < count; i++) indices.push(i);
        for (var j = indices.length - 1; j > 0; j--) {
            var k = Math.floor(Math.random() * (j + 1));
            var swap = indices[j]; indices[j] = indices[k]; indices[k] = swap;
        }
        return indices;
    }

    /* アイテム（または子孫）がテキストを含むか / Whether item (or descendants) contains text */
    function containerHasText(item) {
        if (!item) return false;
        try {
            if (item.typename === "TextFrame") return true;
            if (item.textFrames && item.textFrames.length > 0) return true;
            if (item.pageItems && item.pageItems.length) {
                for (var i = 0; i < item.pageItems.length; i++) {
                    if (containerHasText(item.pageItems[i])) return true;
                }
            }
        } catch (e) { }
        return false;
    }

    /* edittextで上下キーで数値を増減（Shiftで10刻みスナップ） / Up/Down arrow increments edittext value (Shift snaps to 10s) */
    function changeValueByArrowKey(editText, allowNegative, onUpdate) {
        editText.addEventListener("keydown", function (event) {
            if (editText.text.length === 0) return;
            var currentValue = Number(editText.text);
            if (isNaN(currentValue)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            if (event.keyName == "Up" || event.keyName == "Down") {
                var isUp = event.keyName == "Up";
                var delta = 1;
                if (keyboard.shiftKey) {
                    /* 10の倍数にスナップ / Snap to multiples of 10 */
                    currentValue = Math.floor(currentValue / 10) * 10;
                    delta = 10;
                }
                currentValue += isUp ? delta : -delta;
                if (!allowNegative && currentValue < 0) currentValue = 0;

                event.preventDefault();
                editText.text = currentValue;
                if (typeof onUpdate === "function") onUpdate();
            }
        });
    }

    /* N/getLabel/C/R/T/M/B キーで揃えラジオを切替 / Keyboard handler for align radio switching */
    /* 揃えのショートカットキーと、対応するラジオのキー / Shortcut keys mapped to a radio in each row */
    var HORIZONTAL_ALIGN_RADIO_BY_KEY = { "N": "none", "L": "left", "C": "center", "R": "right" };
    var VERTICAL_ALIGN_RADIO_BY_KEY = { "N": "none", "T": "top", "M": "middle", "B": "bottom" };

    /* N / L / C / R / T / M / B キーで揃えのラジオを切り替える / Switch the align radios with the shortcut keys
       alignRadios: { horizontal: {none,left,center,right}, vertical: {none,top,middle,bottom}, getDirection } */
    function addAlignKeyHandler(keyTarget, alignRadios, onUpdate) {
        keyTarget.addEventListener("keydown", function (event) {
            /* 横並びのときは上下の揃え、縦並びのときは左右の揃えを操作する
               A horizontal layout adjusts the vertical align, and vice versa */
            var isHorizontalLayout = (alignRadios.getDirection() === "horizontal");
            var radioByKey = isHorizontalLayout ? VERTICAL_ALIGN_RADIO_BY_KEY : HORIZONTAL_ALIGN_RADIO_BY_KEY;
            var radioSet = isHorizontalLayout ? alignRadios.vertical : alignRadios.horizontal;

            var radioKey = radioByKey[event.keyName];
            if (!radioKey) return;

            for (var key in radioSet) {
                radioSet[key].value = (key === radioKey);
            }
            event.preventDefault();
            if (onUpdate) onUpdate();
        });
    }

    // =========================================
    // アウトライン計測 / Outline measurement
    // =========================================

    /* コンテナ内のすべてのテキストをアウトライン化 / Outline all text in a container */
    function outlineAllTextInContainer(container) {
        if (!container) return;
        try {
            if (container.typename === "TextFrame") {
                try { container.createOutline(); } catch (e0) { }
                return;
            }
            if (container.textFrames && container.textFrames.length) {
                /* createOutline で要素が消えるため後ろから走査 / iterate backwards: createOutline removes the frame */
                for (var i = container.textFrames.length - 1; i >= 0; i--) {
                    try { container.textFrames[i].createOutline(); } catch (e1) { }
                }
            }
        } catch (e) { }
    }

    /* 複製してアウトライン化した境界を計測（boundsCacheに保存して再利用） / Measure outlined bounds once, cached for reuse */
    function measureOutlineBoundsOnce(originalItem, boundsCache) {
        if (!originalItem) return null;

        for (var i = 0; i < boundsCache.length; i++) {
            if (boundsCache[i].item === originalItem) return boundsCache[i].bounds;
        }
        if (!containerHasText(originalItem)) return null;

        var doc = app.activeDocument;
        var measureGroup = null;
        try {
            var layer = null;
            try { layer = originalItem.layer; } catch (eLayer) { }
            if (!layer) layer = doc.activeLayer;
            measureGroup = layer.groupItems.add();
            measureGroup.name = TEMP_MEASURE_GROUP_NAME;

            var duplicatedItem = null;
            try {
                duplicatedItem = originalItem.duplicate(measureGroup, ElementPlacement.PLACEATEND);
            } catch (eDup) {
                duplicatedItem = originalItem.duplicate();
                duplicatedItem.move(measureGroup, ElementPlacement.PLACEATEND);
            }

            outlineAllTextInContainer(duplicatedItem);

            var bounds = null;
            try { bounds = measureGroup.visibleBounds; }
            catch (eVisible) { try { bounds = measureGroup.geometricBounds; } catch (eGeometric) { } }
            if (!bounds) return null;

            boundsCache.push({ item: originalItem, bounds: [bounds[0], bounds[1], bounds[2], bounds[3]] });
            return [bounds[0], bounds[1], bounds[2], bounds[3]];
        } catch (e) {
            return null;
        } finally {
            try { if (measureGroup) measureGroup.remove(); } catch (eRemove) { }
        }
    }

    // =========================================
    // レイアウト適用 / Layout application
    // =========================================

    /* ソート/シャッフルしたアイテムと、ランダム時の基準位置を返す / Return sorted-or-shuffled items + random base position */
    function buildArrangementOrder(items, direction, useRandom, previousShuffleOrder) {
        var arrangement = { items: [], baseLeft: null, baseTop: null, shuffleOrder: previousShuffleOrder };

        if (useRandom) {
            if (!arrangement.shuffleOrder || arrangement.shuffleOrder.length !== items.length) {
                arrangement.shuffleOrder = makeShuffledIndices(items.length);
            }
            for (var si = 0; si < arrangement.shuffleOrder.length; si++) {
                arrangement.items.push(items[arrangement.shuffleOrder[si]]);
            }
            /* ランダム時は元の左上を基準として保持 / Preserve top-left base for random */
            for (var k = 0; k < items.length; k++) {
                var item = items[k];
                if (!item) continue;
                if (arrangement.baseLeft === null || item.left < arrangement.baseLeft) arrangement.baseLeft = item.left;
                if (arrangement.baseTop === null || item.top > arrangement.baseTop) arrangement.baseTop = item.top;
            }
            if (arrangement.baseLeft === null && arrangement.items[0]) arrangement.baseLeft = arrangement.items[0].left;
            if (arrangement.baseTop === null && arrangement.items[0]) arrangement.baseTop = arrangement.items[0].top;
        } else {
            arrangement.shuffleOrder = null;
            arrangement.items = (direction === "horizontal") ? sortByX(items) : sortByY(items);
        }
        return arrangement;
    }

    /* アイテム群の最大幅を計算 / Compute the max item width */
    function computeMaxWidth(items, getBounds) {
        var maxWidth = 0;
        for (var i = 0; i < items.length; i++) {
            var bounds = getBounds(items[i]);
            var width = bounds[2] - bounds[0];
            if (width > maxWidth) maxWidth = width;
        }
        return maxWidth;
    }

    /* アイテム群の最大高さを計算 / Compute the max item height */
    function computeMaxHeight(items, getBounds) {
        var maxHeight = 0;
        for (var i = 0; i < items.length; i++) {
            var bounds = getBounds(items[i]);
            var height = bounds[1] - bounds[3];
            if (height > maxHeight) maxHeight = height;
        }
        return maxHeight;
    }

    /* 横並び配置（左→右、上下揃えを適用） / Horizontal placement (left → right, with vertical align) */
    function applyHorizontalLayout(items, startX, startY, referenceHeight, spacingPt, vAlignMode, getBounds) {
        var currentX = startX;
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) continue;
            var bounds = getBounds(item);
            var itemWidth = bounds[2] - bounds[0];

            /* Xは左基準で配置 / Place by left edge on X */
            item.left = item.left + (currentX - bounds[0]);

            /* 上下揃え / Vertical alignment */
            if (vAlignMode !== "none") {
                var cellTop = startY;
                var cellBottom = startY - referenceHeight;
                var dy = 0;
                if (vAlignMode === "middle") dy = ((cellTop + cellBottom) / 2) - ((bounds[1] + bounds[3]) / 2);
                else if (vAlignMode === "bottom") dy = cellBottom - bounds[3];
                else dy = cellTop - bounds[1]; /* top */
                item.top = item.top + dy;
            }

            if (i < items.length - 1) {
                currentX += itemWidth + spacingPt;
            }
        }
    }

    /* 縦並び配置（上→下、左右揃えを適用） / Vertical placement (top → bottom, with horizontal align) */
    function applyVerticalLayout(items, startX, startY, referenceWidth, spacingPt, hAlignMode, getBounds) {
        var currentY = startY;
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) continue;
            var bounds = getBounds(item);
            var itemHeight = bounds[1] - bounds[3];

            /* 左右揃え / Horizontal alignment */
            if (hAlignMode !== "none") {
                var cellLeft = startX;
                var cellRight = cellLeft + referenceWidth;
                var dx = 0;
                if (hAlignMode === "center") dx = ((cellLeft + cellRight) / 2) - ((bounds[0] + bounds[2]) / 2);
                else if (hAlignMode === "right") dx = cellRight - bounds[2];
                else dx = cellLeft - bounds[0]; /* left */
                item.left = item.left + dx;
            }

            /* YはTop揃えで積む / Stack by top on Y */
            item.top = item.top + (currentY - bounds[1]);

            if (i < items.length - 1) {
                currentY -= itemHeight + spacingPt;
            }
        }
    }

    /* ランダム時に元の左上位置へオフセット補正 / Offset items so the first stays at original top-left (random) */
    function applyRandomBaseOffset(items, baseLeft, baseTop) {
        if (!items.length) return;
        var dx = baseLeft - items[0].left;
        var dy = baseTop - items[0].top;
        for (var i = 0; i < items.length; i++) {
            if (!items[i]) continue;
            items[i].left += dx;
            items[i].top += dy;
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* メインダイアログを構築して整列/分布を実行 / Build the main dialog and run align/distribute */
    function showArrangeDialog() {
        var alignDialog = new Window("dialog", getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        alignDialog.orientation = "column";
        alignDialog.alignChildren = "fill";
        alignDialog.opacity = DIALOG_OPACITY;

        var lastPosition = loadDialogPosition();
        if (lastPosition) alignDialog.location = lastPosition;

        /* キャンセル時に復元するため現在のプリファレンスを保存 / Preserve preference for cancel */
        var originalIncludeStrokeInBounds = app.preferences.getBooleanPreference("includeStrokeInBounds");

        /* 選択スナップショット / Snapshot selection */
        var originalSelection = app.activeDocument.selection.slice();
        var previewState = { isUndo: false };
        var detectedDirection = detectDirection(originalSelection);

        /* 選択にテキストが含まれているか / Whether selection contains any text */
        var selectionHasText = false;
        for (var si = 0; si < originalSelection.length; si++) {
            if (containerHasText(originalSelection[si])) { selectionHasText = true; break; }
        }

        /* ---- UI: 方向 / Direction ---- */
        var directionPanel = alignDialog.add("panel", undefined, getLabel('direction.panel'));
        directionPanel.orientation = "row";
        directionPanel.alignChildren = ["center", "center"];
        directionPanel.margins = [15, 20, 15, 10];
        var rbDirAuto = directionPanel.add("radiobutton", undefined, getLabel('direction.auto'));
        rbDirAuto.helpTip = getLabel('direction.autoTip');
        var rbDirVertical = directionPanel.add("radiobutton", undefined, getLabel('direction.vertical'));
        rbDirVertical.helpTip = getLabel('direction.verticalTip');
        var rbDirHorizontal = directionPanel.add("radiobutton", undefined, getLabel('direction.horizontal'));
        rbDirHorizontal.helpTip = getLabel('direction.horizontalTip');
        rbDirVertical.value = true; /* デフォルトを「縦」に / Default to Vertical */

        /* 選択中のラジオから実効方向を返す（自動は判定結果） / Resolve effective direction from radios (auto → detected) */
        function getEffectiveDirection() {
            if (rbDirVertical.value) return "vertical";
            if (rbDirHorizontal.value) return "horizontal";
            return detectedDirection;
        }

        /* ---- UI: 間隔 / Spacing ---- */
        var spacingPanel = alignDialog.add("panel", undefined, getLabel('spacing.label'));
        spacingPanel.orientation = "column";
        spacingPanel.alignChildren = ["center", "center"];
        spacingPanel.margins = [15, 20, 15, 10];
        var spacingRowGroup = spacingPanel.add("group");
        spacingRowGroup.orientation = "row";
        spacingRowGroup.alignChildren = ["left", "center"];
        var spacingInput = spacingRowGroup.add("edittext", undefined, "0");
        spacingInput.characters = 3;
        spacingInput.helpTip = getLabel('spacing.tip');
        spacingRowGroup.add("statictext", undefined, getUnitInfo().label);
        changeValueByArrowKey(spacingInput, true, function () { requestPreviewUpdate(); });

        /* ---- UI: 揃え / Alignment ---- */
        var alignPanel = alignDialog.add("panel", undefined, "");
        alignPanel.orientation = "column";
        alignPanel.alignChildren = ["left", "center"];
        alignPanel.margins = [15, 20, 15, 10];

        var hAlignGroup = alignPanel.add("group");
        hAlignGroup.orientation = "row";
        hAlignGroup.alignChildren = ["left", "center"];
        var rbHNone = hAlignGroup.add("radiobutton", undefined, getLabel('align.hNone'));
        var rbHLeft = hAlignGroup.add("radiobutton", undefined, getLabel('align.hLeft'));
        var rbHCenter = hAlignGroup.add("radiobutton", undefined, getLabel('align.hCenter'));
        var rbHRight = hAlignGroup.add("radiobutton", undefined, getLabel('align.hRight'));
        rbHCenter.value = true; /* デフォルトを「中央」に / Default to Center */
        rbHNone.helpTip = rbHLeft.helpTip = rbHCenter.helpTip = rbHRight.helpTip = getLabel('align.horizontalTip');

        var vAlignGroup = alignPanel.add("group");
        vAlignGroup.orientation = "row";
        vAlignGroup.alignChildren = ["left", "center"];
        var rbVNone = vAlignGroup.add("radiobutton", undefined, getLabel('align.vNone'));
        var rbVTop = vAlignGroup.add("radiobutton", undefined, getLabel('align.vTop'));
        var rbVMiddle = vAlignGroup.add("radiobutton", undefined, getLabel('align.vMiddle'));
        var rbVBottom = vAlignGroup.add("radiobutton", undefined, getLabel('align.vBottom'));
        rbVMiddle.value = true;
        rbVNone.helpTip = rbVTop.helpTip = rbVMiddle.helpTip = rbVBottom.helpTip = getLabel('align.verticalTip');

        /* キーハンドラーへ渡すラジオ一式 / The radio set handed to the key handler */
        var alignRadios = {
            horizontal: { none: rbHNone, left: rbHLeft, center: rbHCenter, right: rbHRight },
            vertical: { none: rbVNone, top: rbVTop, middle: rbVMiddle, bottom: rbVBottom },
            getDirection: getEffectiveDirection
        };

        rbVNone.onClick = function () { requestPreviewUpdate(); };
        rbVTop.onClick = function () { requestPreviewUpdate(); };
        rbVMiddle.onClick = function () { requestPreviewUpdate(); };
        rbVBottom.onClick = function () { requestPreviewUpdate(); };
        rbHLeft.onClick = function () { requestPreviewUpdate(); };
        rbHCenter.onClick = function () { requestPreviewUpdate(); };
        rbHRight.onClick = function () { requestPreviewUpdate(); };
        rbHNone.onClick = function () { requestPreviewUpdate(); };

        /* ---- UI: オプション / Options ---- */
        var optionsGroup = alignDialog.add("group");
        optionsGroup.orientation = "column";
        optionsGroup.alignChildren = ["left", "center"];
        optionsGroup.alignment = ["fill", "top"];
        optionsGroup.margins = [15, 5, 15, 5];

        var usePreviewBoundsCheckbox = optionsGroup.add("checkbox", undefined, getLabel('options.useBounds'));
        usePreviewBoundsCheckbox.value = true;
        usePreviewBoundsCheckbox.helpTip = getLabel('options.useBoundsTip');
        usePreviewBoundsCheckbox.onClick = function () {
            /* 境界モードが変わるとアウトライン計測結果も再計算 / Outline cache invalid when bounds mode changes */
            outlineMeasureCache = [];
            requestPreviewUpdate();
        };

        var measureTextOutlineCheckbox = optionsGroup.add("checkbox", undefined, getLabel('options.measureText'));
        measureTextOutlineCheckbox.value = false;
        measureTextOutlineCheckbox.helpTip = getLabel('options.measureTextTip');

        var randomCheckbox = optionsGroup.add("checkbox", undefined, getLabel('options.random'));
        randomCheckbox.value = false;
        randomCheckbox.helpTip = getLabel('options.randomTip');

        /* ---- UI: ボタン / Buttons ---- */
        var btnRowGroup = alignDialog.add("group");
        btnRowGroup.alignment = "center";
        btnRowGroup.alignChildren = ["center", "center"];
        btnRowGroup.add("button", undefined, getLabel('common.cancel'), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel('common.ok'), { name: "ok" });

        /* ---- 内部状態 / Internal state ---- */
        var randomOrderCache = null;
        var outlineMeasureCache = []; /* [{item:PageItem, bounds:[l,t,r,b]}] */

        /* レイアウト用に境界を取得（縦並び時はテキスト計測キャッシュ経由） / Get bounds for layout (outline cache when vertical) */
        function getBoundsForLayout(item) {
            var direction = getEffectiveDirection();
            if (measureTextOutlineCheckbox.value && direction !== "horizontal") {
                var outlinedBounds = measureOutlineBoundsOnce(item, outlineMeasureCache);
                if (outlinedBounds) return outlinedBounds;
            }
            return getItemBounds(item, usePreviewBoundsCheckbox.value);
        }

        /* 左右揃えモードを取得 / Get horizontal align mode */
        function getHAlignMode() {
            if (rbHNone.value) return "none";
            if (rbHCenter.value) return "center";
            if (rbHRight.value) return "right";
            return "left";
        }

        /* 上下揃えモードを取得 / Get vertical align mode */
        function getVAlignMode() {
            if (rbVNone.value) return "none";
            if (rbVMiddle.value) return "middle";
            if (rbVBottom.value) return "bottom";
            return "top";
        }

        /* 入力値を現在の単位からポイントへ換算 / Convert input value from current unit to points */
        function getSpacingPt() {
            var spacingValue = parseFloat(spacingInput.text);
            if (isNaN(spacingValue)) spacingValue = 0;
            return spacingValue * getUnitInfo().pointsPerUnit;
        }

        /* 現在の設定で選択にレイアウトを適用 / Apply layout to current selection with current settings */
        function applyLayoutToSelection() {
            if (!originalSelection || originalSelection.length === 0) return;

            var direction = getEffectiveDirection();
            var spacingPt = getSpacingPt();

            var arrangement = buildArrangementOrder(originalSelection, direction, randomCheckbox.value, randomOrderCache);
            randomOrderCache = arrangement.shuffleOrder;
            var items = arrangement.items;
            if (!items.length) return;

            var startBounds = getBoundsForLayout(items[0]);
            var startX = startBounds[0];
            var startY = startBounds[1];

            if (direction === "horizontal") {
                var referenceHeight = computeMaxHeight(items, getBoundsForLayout);
                applyHorizontalLayout(items, startX, startY, referenceHeight, spacingPt, getVAlignMode(), getBoundsForLayout);
            } else {
                var referenceWidth = computeMaxWidth(items, getBoundsForLayout);
                applyVerticalLayout(items, startX, startY, referenceWidth, spacingPt, getHAlignMode(), getBoundsForLayout);
            }

            if (randomCheckbox.value) {
                applyRandomBaseOffset(items, arrangement.baseLeft, arrangement.baseTop);
            }
        }

        /* 方向に応じて揃えパネル/オプションの有効状態を同期 / Sync enabled state of align panel & options to direction */
        function syncAlignUI() {
            var direction = getEffectiveDirection();
            hAlignGroup.visible = true;
            vAlignGroup.visible = true;

            if (direction === "horizontal") {
                alignPanel.text = getLabel('align.verticalTitle');
                hAlignGroup.enabled = false;
                vAlignGroup.enabled = true;
            } else {
                alignPanel.text = getLabel('align.horizontalTitle');
                hAlignGroup.enabled = true;
                vAlignGroup.enabled = false;
            }

            /* テキスト計測は縦並び時かつ選択にテキストがある場合のみ有効 / Enable text-measure only when vertical & selection has text */
            try {
                measureTextOutlineCheckbox.enabled = (direction !== "horizontal") && selectionHasText;
            } catch (e) { }

            try { alignDialog.layout.layout(true); } catch (e) { }
        }

        /* ---- プレビュー更新（前回 preview を app.undo() で巻き戻して再生成） / Preview update via app.undo() ---- */
        var __isPreviewUpdating = false;
        var __lastPreviewAt = 0;

        /* 連続呼び出しを間引きつつプレビューを再描画 / Re-render preview, throttled against rapid calls */
        function updatePreview() {
            if (__isPreviewUpdating) return;
            var now = new Date().getTime();
            if (now - __lastPreviewAt < PREVIEW_MIN_INTERVAL_MS) return;
            __lastPreviewAt = now;

            __isPreviewUpdating = true;
            try {
                try {
                    app.preferences.setBooleanPreference("includeStrokeInBounds", usePreviewBoundsCheckbox.value);
                } catch (e) { }
                runPreview(previewState, applyLayoutToSelection, true);
            } finally {
                __isPreviewUpdating = false;
            }
        }

        /* ScriptUIイベント中のクラッシュ回避のためscheduleTaskで遅延実行 / Defer to avoid crash inside ScriptUI events */
        var __previewTaskId = 0;
        $.global.__SAT_updatePreview = updatePreview;
        function requestPreviewUpdate() {
            try { if (__previewTaskId) app.cancelTask(__previewTaskId); } catch (e) { }
            $.global.__SAT_updatePreview = updatePreview;
            try {
                __previewTaskId = app.scheduleTask('$.global.__SAT_updatePreview && $.global.__SAT_updatePreview();', PREVIEW_SCHEDULE_MS, false);
            } catch (e) {
                updatePreview();
            }
        }

        /* ---- イベントバインド / Event bindings ---- */
        addAlignKeyHandler(alignDialog, alignRadios, requestPreviewUpdate);
        addAlignKeyHandler(spacingInput, alignRadios, requestPreviewUpdate);

        /* 方向変更時に揃えUIを同期してプレビュー更新 / On direction change, sync align UI and refresh preview */
        function onDirChanged() { syncAlignUI(); requestPreviewUpdate(); }
        rbDirAuto.onClick = onDirChanged;
        rbDirVertical.onClick = onDirChanged;
        rbDirHorizontal.onClick = onDirChanged;

        randomCheckbox.onClick = function () {
            randomOrderCache = null;
            requestPreviewUpdate();
        };

        measureTextOutlineCheckbox.onClick = function () {
            outlineMeasureCache = [];
            requestPreviewUpdate();
        };

        /* 初期化 / Init */
        syncAlignUI();
        requestPreviewUpdate();
        spacingInput.active = true;

        var dialogResult = alignDialog.show();
        try { if (__previewTaskId) app.cancelTask(__previewTaskId); } catch (e) { }
        saveDialogPosition(alignDialog.location);

        if (dialogResult !== 1) {
            /* キャンセル: プレビューを undo してプリファレンスも戻す / Cancel: undo preview & restore preference */
            cleanupPreview(previewState, app.activeDocument);
            app.preferences.setBooleanPreference("includeStrokeInBounds", originalIncludeStrokeInBounds);
            app.redraw();
            return false;
        }

        /* 確定: プレビュー分を巻き戻して本実行（undo 履歴を 1 件にまとめる） / OK: undo preview, then commit as single history entry */
        undoPreview(previewState);
        try { app.preferences.setBooleanPreference("includeStrokeInBounds", usePreviewBoundsCheckbox.value); } catch (e) { }
        applyLayoutToSelection();
        app.redraw();
        return true;
    }

    // =========================================
    // メイン / Main
    // =========================================

    /* 選択を検証してダイアログを起動 / Validate selection and launch the dialog */
    function main() {
        try {
            var selectedItems = app.activeDocument.selection;
            if (!selectedItems || selectedItems.length === 0) {
                alert(getLabel('error.needSelection'));
                return;
            }
            showArrangeDialog();
        } catch (e) {
            alert(getLabel('error.prefix') + e.message);
        }
    }

    main();

})();
