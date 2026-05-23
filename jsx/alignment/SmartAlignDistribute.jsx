#target illustrator
#targetengine "SmartAlignAndTileEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名
SmartAlignDistribute.jsx

### 概要 / Overview
選択オブジェクトを「縦/横」に並べて、指定した間隔で分布します。方向は自動判定も可能で、揃え（左右/上下）、プレビュー境界（visible/geometric）、ランダム並べ替えにも対応します。
縦並び時にテキストを含む場合、ダイアログ中だけ一度だけ計測用に複製→アウトライン化して高さ（必要に応じて幅も）を計測し、その結果をダイアログ中だけキャッシュします（プレビューのたびに複製しない）。

### オリジナルアイデア
John Wundes
Distribute Stacked Objects v1.1
https://github.com/johnwun/js4ai/blob/master/distributeStackedObjects.jsx

Gorolib Design
https://gorolib.blog.jp/archives/77282974.html

### 更新日 / Updated: 2026-02-26
*/

(function () {

    // =========================================
    // バージョン / Version
    // =========================================
    var SCRIPT_VERSION = "v1.2.0";

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var dialogOpacity = 0.97;
    var DLG_POS_MEM_KEY = "__SmartAlignAndTileTateSimple_DlgPos__";
    var PREVIEW_MIN_INTERVAL_MS = 80;
    var PREVIEW_SCHEDULE_MS = 60;

    // =========================================
    // ローカライズ / Localization
    // =========================================
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    var LABELS = {
        /* === 共通 / Common === */
        dialogTitle: { ja: "整列と分布", en: "Align & Distribute" },
        cancel: { ja: "キャンセル", en: "Cancel" },

        /* === 方向 / Direction === */
        directionPanel: { ja: "方向", en: "Direction" },
        dirAuto: { ja: "自動", en: "Auto" },
        dirVertical: { ja: "縦", en: "Vertical" },
        dirHorizontal: { ja: "横", en: "Horizontal" },

        /* === 間隔 / Spacing === */
        spacingLabel: { ja: "間隔", en: "Spacing" },

        /* === 揃え / Alignment === */
        alignH: { ja: "揃え（左右）", en: "Align (H)" },
        alignV: { ja: "揃え（上下）", en: "Align (V)" },
        hAlignLeft: { ja: "左", en: "Left" },
        hAlignCenter: { ja: "中央", en: "Center" },
        hAlignRight: { ja: "右", en: "Right" },
        hAlignNone: { ja: "なし", en: "None" },
        vAlignTop: { ja: "上", en: "Top" },
        vAlignMiddle: { ja: "中央", en: "Middle" },
        vAlignBottom: { ja: "下", en: "Bottom" },
        vAlignNone: { ja: "なし", en: "None" },

        /* === オプション / Options === */
        useBounds: { ja: "プレビュー境界を使用", en: "Use preview bounds" },
        random: { ja: "ランダム", en: "Random" },
        measureTextOutline: { ja: "テキストの高さを計測", en: "V: Measure text" },
        measureTextOutlineTip: {
            ja: "縦並び時のみ、テキストを一度だけ複製→アウトライン化して境界を計測します（ダイアログ中だけキャッシュ）。",
            en: "Only in vertical layout, measures text by duplicating and outlining once (cached for this dialog only)."
        },

        /* === エラー / Errors === */
        needSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
        errorPrefix: { ja: "エラーが発生しました: ", en: "An error has occurred: " }
    };

    /* キーで LABELS から現在言語の文字列を取得 / Get current-language string from LABELS by key */
    function L(key) {
        return LABELS[key][lang];
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(key) {
        return L(key) + (lang === 'ja' ? '：' : ':');
    }

    /* 件数付きラベル（日本語は全角括弧、英語は半角括弧）/ Label with count (full-width JA, half-width EN) */
    function labelWithCount(key, count) {
        if (lang === "ja") return L(key) + "（" + count + "）";
        return L(key) + " (" + count + ")";
    }

    // =========================================
    // 単位 / Units
    // =========================================
    var unitLabelMap = {
        0: "in", 1: "mm", 2: "pt", 3: "pica", 4: "cm",
        5: "Q/H", 6: "px", 7: "ft/in", 8: "m", 9: "yd", 10: "ft"
    };

    /* 現在の単位ラベルを取得 / Get current unit label */
    function getCurrentUnitLabel() {
        var unitCode = app.preferences.getIntegerPreference("rulerType");
        return unitLabelMap[unitCode] || "pt";
    }

    /* 単位コードからポイント換算係数を取得 / Get point factor from unit code */
    function getPtFactorFromUnitCode(code) {
        switch (code) {
            case 0: return 72.0;                    // in
            case 1: return 72.0 / 25.4;             // mm
            case 2: return 1.0;                     // pt
            case 3: return 12.0;                    // pica
            case 4: return 72.0 / 2.54;             // cm
            case 5: return 72.0 / 25.4 * 0.25;      // Q / H
            case 6: return 1.0;                     // px
            case 7: return 72.0 * 12.0;             // ft/in
            case 8: return 72.0 / 25.4 * 1000.0;    // m
            case 9: return 72.0 * 36.0;             // yd
            case 10: return 72.0 * 12.0;             // ft
            default: return 1.0;
        }
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

    /* ダイアログクローズ時のクリーンアップ / Cleanup on dialog close */
    function cleanupPreview(state, doc, tempLayerName) {
        try {
            if (state.isUndo) app.undo();
            state.isUndo = false;
        } catch (err) { }
        if (tempLayerName) {
            try {
                var tmpLay = doc.layers.getByName(tempLayerName);
                tmpLay.remove();
            } catch (err) { }
        }
    }

    /* ダイアログ位置をセッション内に保存/復元 / Save & load dialog position within session */
    function loadDialogPosition() {
        try {
            var p = $.global[DLG_POS_MEM_KEY];
            if (p && p.length === 2) return [p[0], p[1]];
        } catch (e) { }
        return null;
    }
    function saveDialogPosition(loc) {
        try {
            if (!loc || loc.length !== 2) return;
            $.global[DLG_POS_MEM_KEY] = [Math.round(loc[0]), Math.round(loc[1])];
        } catch (e) { }
    }

    /* アイテムの境界を取得（visible/geometric） / Get item bounds */
    function getItemBounds(item, usePreviewBounds) {
        return usePreviewBounds ? item.visibleBounds : item.geometricBounds;
    }

    /* 選択範囲の幅・高さスパンを取得 / Get the span of a selection */
    function getSelectionSpan(items, usePreviewBounds) {
        var minL = null, maxR = null, maxT = null, minB = null;
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (!it) continue;
            var b = getItemBounds(it, usePreviewBounds);
            if (minL === null || b[0] < minL) minL = b[0];
            if (maxR === null || b[2] > maxR) maxR = b[2];
            if (maxT === null || b[1] > maxT) maxT = b[1];
            if (minB === null || b[3] < minB) minB = b[3];
        }
        if (minL === null) return { spanX: 0, spanY: 0 };
        return { spanX: (maxR - minL), spanY: (maxT - minB) };
    }

    /* 自動方向判定（geometricBoundsで安定判定） / Auto-detect direction */
    function detectDirection(items) {
        var s = getSelectionSpan(items, false);
        return (s.spanX >= s.spanY) ? "horizontal" : "vertical";
    }

    /* Y座標でソート（上→下、同じなら左→右） / Sort by Y (top to bottom, then left) */
    function sortByY(items) {
        var copy = items.slice();
        copy.sort(function (a, b) {
            if (a.top !== b.top) return b.top - a.top;
            return a.left - b.left;
        });
        return copy;
    }

    /* X座標でソート（左→右、同じなら上→下） / Sort by X (left to right, then top) */
    function sortByX(items) {
        var copy = items.slice();
        copy.sort(function (a, b) {
            if (a.left !== b.left) return a.left - b.left;
            return b.top - a.top;
        });
        return copy;
    }

    /* Fisher-Yatesでインデックス配列をシャッフル / Fisher-Yates shuffle of index array */
    function makeShuffledIndices(length) {
        var order = [];
        for (var i = 0; i < length; i++) order.push(i);
        for (var j = order.length - 1; j > 0; j--) {
            var k = Math.floor(Math.random() * (j + 1));
            var tmp = order[j]; order[j] = order[k]; order[k] = tmp;
        }
        return order;
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

    /* edittextで上下キーで数値を増減（Shiftで10刻みスナップ） / Up/Down arrow increments edittext value */
    function changeValueByArrowKey(editText, allowNegative, onUpdate) {
        editText.addEventListener("keydown", function (event) {
            if (editText.text.length === 0) return;
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            if (event.keyName == "Up" || event.keyName == "Down") {
                var isUp = event.keyName == "Up";
                var delta = 1;
                if (keyboard.shiftKey) {
                    /* 10の倍数にスナップ / Snap to multiples of 10 */
                    value = Math.floor(value / 10) * 10;
                    delta = 10;
                }
                value += isUp ? delta : -delta;
                if (!allowNegative && value < 0) value = 0;

                event.preventDefault();
                editText.text = value;
                if (typeof onUpdate === "function") onUpdate();
            }
        });
    }

    /* N/L/C/R/T/M/B キーで揃えラジオを切替 / Keyboard handler for align radio switching */
    function addAlignKeyHandler(target, getDir, rbHNone, rbHLeft, rbHCenter, rbHRight, rbVNone, rbVTop, rbVMiddle, rbVBottom, onUpdate) {
        target.addEventListener("keydown", function (event) {
            var k = event.keyName;
            var dir = (typeof getDir === "function") ? getDir() : "vertical";

            function applyAndUpdate() {
                event.preventDefault();
                if (typeof onUpdate === "function") onUpdate();
            }

            /* "なし" は両方向共通 / "None" works for both axes */
            if (k === "N") {
                if (dir === "horizontal") rbVNone.value = true;
                else rbHNone.value = true;
                applyAndUpdate();
                return;
            }

            /* 縦並び: L/C/R で左右揃え / Vertical stacking: L/C/R for horizontal alignment */
            if (dir !== "horizontal") {
                if (k === "L") { rbHLeft.value = true; applyAndUpdate(); return; }
                if (k === "C") { rbHCenter.value = true; applyAndUpdate(); return; }
                if (k === "R") { rbHRight.value = true; applyAndUpdate(); return; }
                return;
            }

            /* 横並び: T/M/B で上下揃え / Horizontal stacking: T/M/B for vertical alignment */
            if (k === "T") { rbVTop.value = true; applyAndUpdate(); return; }
            if (k === "M") { rbVMiddle.value = true; applyAndUpdate(); return; }
            if (k === "B") { rbVBottom.value = true; applyAndUpdate(); return; }
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

    /* 複製してアウトライン化した境界を計測（cacheに保存） / Measure outlined bounds with cache */
    function measureOutlineBoundsOnce(originalItem, cache) {
        if (!originalItem) return null;

        for (var i = 0; i < cache.length; i++) {
            if (cache[i].ref === originalItem) return cache[i].b;
        }
        if (!containerHasText(originalItem)) return null;

        var doc = activeDocument;
        var tempGroup = null;
        try {
            var layer = null;
            try { layer = originalItem.layer; } catch (eLayer) { }
            if (!layer) layer = doc.activeLayer;
            tempGroup = layer.groupItems.add();
            try { tempGroup.name = "__SAT_TempMeasure__"; } catch (_) { }

            var dup = null;
            try {
                dup = originalItem.duplicate(tempGroup, ElementPlacement.PLACEATEND);
            } catch (eDup) {
                dup = originalItem.duplicate();
                try { dup.move(tempGroup, ElementPlacement.PLACEATEND); } catch (_) { }
            }

            outlineAllTextInContainer(dup);

            var b = null;
            try { b = tempGroup.visibleBounds; }
            catch (eB) { try { b = tempGroup.geometricBounds; } catch (eB2) { } }
            if (!b) return null;

            cache.push({ ref: originalItem, b: [b[0], b[1], b[2], b[3]] });
            return [b[0], b[1], b[2], b[3]];
        } catch (e) {
            return null;
        } finally {
            try { if (tempGroup) tempGroup.remove(); } catch (eRm) { }
        }
    }

    // =========================================
    // レイアウト適用 / Layout application
    // =========================================

    /* ソート/シャッフルしたアイテムと、ランダム時の基準位置 / Sorted-or-shuffled items + random base */
    function buildArrangementOrder(items, direction, useRandom, prevRandomCache) {
        var result = { items: [], baseLeft: null, baseTop: null, cache: prevRandomCache };

        if (useRandom) {
            if (!result.cache || result.cache.length !== items.length) {
                result.cache = makeShuffledIndices(items.length);
            }
            for (var si = 0; si < result.cache.length; si++) {
                result.items.push(items[result.cache[si]]);
            }
            /* ランダム時は元の左上を基準として保持 / Preserve top-left base for random */
            for (var k = 0; k < items.length; k++) {
                var it = items[k];
                if (!it) continue;
                if (result.baseLeft === null || it.left < result.baseLeft) result.baseLeft = it.left;
                if (result.baseTop === null || it.top > result.baseTop) result.baseTop = it.top;
            }
            if (result.baseLeft === null && result.items[0]) result.baseLeft = result.items[0].left;
            if (result.baseTop === null && result.items[0]) result.baseTop = result.items[0].top;
        } else {
            result.cache = null;
            result.items = (direction === "horizontal") ? sortByX(items) : sortByY(items);
        }
        return result;
    }

    /* 最大幅 / Max item width */
    function computeMaxWidth(items, getBounds) {
        var max = 0;
        for (var i = 0; i < items.length; i++) {
            var b = getBounds(items[i]);
            var w = b[2] - b[0];
            if (w > max) max = w;
        }
        return max;
    }

    /* 最大高さ / Max item height */
    function computeMaxHeight(items, getBounds) {
        var max = 0;
        for (var i = 0; i < items.length; i++) {
            var b = getBounds(items[i]);
            var h = b[1] - b[3];
            if (h > max) max = h;
        }
        return max;
    }

    /* 横並び配置（左→右） / Horizontal placement (left → right) */
    function applyHorizontalLayout(items, startX, startY, refHeight, spacingPt, vAlignMode, getBounds) {
        var currentX = startX;
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) continue;
            var b = getBounds(item);
            var itemWidth = b[2] - b[0];

            /* Xは左基準で配置 / Place by left edge on X */
            item.left = item.left + (currentX - b[0]);

            /* 上下揃え / Vertical alignment */
            if (vAlignMode !== "none") {
                var cellTop = startY;
                var cellBottom = startY - refHeight;
                var dy = 0;
                if (vAlignMode === "middle") dy = ((cellTop + cellBottom) / 2) - ((b[1] + b[3]) / 2);
                else if (vAlignMode === "bottom") dy = cellBottom - b[3];
                else dy = cellTop - b[1]; /* top */
                item.top = item.top + dy;
            }

            if (i < items.length - 1) {
                currentX += itemWidth + spacingPt;
            }
        }
    }

    /* 縦並び配置（上→下） / Vertical placement (top → bottom) */
    function applyVerticalLayout(items, startX, startY, refWidth, spacingPt, hAlignMode, getBounds) {
        var currentY = startY;
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) continue;
            var b = getBounds(item);
            var itemHeight = b[1] - b[3];

            /* 左右揃え / Horizontal alignment */
            if (hAlignMode !== "none") {
                var cellLeft = startX;
                var cellRight = cellLeft + refWidth;
                var dx = 0;
                if (hAlignMode === "center") dx = ((cellLeft + cellRight) / 2) - ((b[0] + b[2]) / 2);
                else if (hAlignMode === "right") dx = cellRight - b[2];
                else dx = cellLeft - b[0]; /* left */
                item.left = item.left + dx;
            }

            /* YはTop揃えで積む / Stack by top on Y */
            item.top = item.top + (currentY - b[1]);

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

    function showArrangeDialog() {
        var dlg = new Window("dialog", L('dialogTitle') + ' ' + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = "fill";
        dlg.opacity = dialogOpacity;

        var lastPos = loadDialogPosition();
        if (lastPos) dlg.location = lastPos;

        /* キャンセル時に復元するため現在のプリファレンスを保存 / Preserve preference for cancel */
        var originalIncludeStrokeInBounds = app.preferences.getBooleanPreference("includeStrokeInBounds");

        /* 選択スナップショット / Snapshot selection */
        var originalSelection = activeDocument.selection.slice();
        var previewState = { isUndo: false };
        var detectedDir = detectDirection(originalSelection);

        /* 選択にテキストが含まれているか / Whether selection contains any text */
        var selectionHasText = false;
        for (var si = 0; si < originalSelection.length; si++) {
            if (containerHasText(originalSelection[si])) { selectionHasText = true; break; }
        }

        /* ---- UI: 方向 / Direction ---- */
        var directionPanel = dlg.add("panel", undefined, L('directionPanel'));
        directionPanel.orientation = "row";
        directionPanel.alignChildren = ["center", "center"];
        directionPanel.margins = [15, 20, 15, 10];
        var rbDirAuto = directionPanel.add("radiobutton", undefined, L('dirAuto'));
        var rbDirV = directionPanel.add("radiobutton", undefined, L('dirVertical'));
        var rbDirH = directionPanel.add("radiobutton", undefined, L('dirHorizontal'));
        rbDirV.value = true; /* デフォルトを「縦」に / Default to Vertical */

        function getEffectiveDirection() {
            if (rbDirV.value) return "vertical";
            if (rbDirH.value) return "horizontal";
            return detectedDir;
        }

        /* ---- UI: 間隔 / Spacing ---- */
        var spacingPanel = dlg.add("panel", undefined, L('spacingLabel'));
        spacingPanel.orientation = "column";
        spacingPanel.alignChildren = ["center", "center"];
        spacingPanel.margins = [15, 20, 15, 10];
        var spacingRow = spacingPanel.add("group");
        spacingRow.orientation = "row";
        spacingRow.alignChildren = ["left", "center"];
        var spacingInput = spacingRow.add("edittext", undefined, "0");
        spacingInput.characters = 3;
        spacingRow.add("statictext", undefined, getCurrentUnitLabel());
        changeValueByArrowKey(spacingInput, true, function () { requestPreviewUpdate(); });

        /* ---- UI: 揃え / Alignment ---- */
        var alignPanel = dlg.add("panel", undefined, "");
        alignPanel.orientation = "column";
        alignPanel.alignChildren = ["left", "center"];
        alignPanel.margins = [15, 20, 15, 10];

        var hAlignGroup = alignPanel.add("group");
        hAlignGroup.orientation = "row";
        hAlignGroup.alignChildren = ["left", "center"];
        var rbHNone = hAlignGroup.add("radiobutton", undefined, L('hAlignNone'));
        var rbHLeft = hAlignGroup.add("radiobutton", undefined, L('hAlignLeft'));
        var rbHCenter = hAlignGroup.add("radiobutton", undefined, L('hAlignCenter'));
        var rbHRight = hAlignGroup.add("radiobutton", undefined, L('hAlignRight'));
        rbHCenter.value = true; /* デフォルトを「中央」に / Default to Center */

        var vAlignGroup = alignPanel.add("group");
        vAlignGroup.orientation = "row";
        vAlignGroup.alignChildren = ["left", "center"];
        var rbVNone = vAlignGroup.add("radiobutton", undefined, L('vAlignNone'));
        var rbVTop = vAlignGroup.add("radiobutton", undefined, L('vAlignTop'));
        var rbVMiddle = vAlignGroup.add("radiobutton", undefined, L('vAlignMiddle'));
        var rbVBottom = vAlignGroup.add("radiobutton", undefined, L('vAlignBottom'));
        rbVMiddle.value = true;

        rbVNone.onClick = function () { requestPreviewUpdate(); };
        rbVTop.onClick = function () { requestPreviewUpdate(); };
        rbVMiddle.onClick = function () { requestPreviewUpdate(); };
        rbVBottom.onClick = function () { requestPreviewUpdate(); };
        rbHLeft.onClick = function () { requestPreviewUpdate(); };
        rbHCenter.onClick = function () { requestPreviewUpdate(); };
        rbHRight.onClick = function () { requestPreviewUpdate(); };
        rbHNone.onClick = function () { requestPreviewUpdate(); };

        /* ---- UI: オプション / Options ---- */
        var optionsGroup = dlg.add("group");
        optionsGroup.orientation = "column";
        optionsGroup.alignChildren = ["left", "center"];
        optionsGroup.alignment = ["fill", "top"];
        optionsGroup.margins = [15, 5, 15, 5];

        var usePreviewBoundsCheckbox = optionsGroup.add("checkbox", undefined, L('useBounds'));
        usePreviewBoundsCheckbox.value = true;
        usePreviewBoundsCheckbox.onClick = function () {
            /* 境界モードが変わるとアウトライン計測結果も再計算 / Outline cache invalid when bounds mode changes */
            outlineMeasureCache = [];
            requestPreviewUpdate();
        };

        var measureTextOutlineCheckbox = optionsGroup.add("checkbox", undefined, L('measureTextOutline'));
        measureTextOutlineCheckbox.value = false;
        measureTextOutlineCheckbox.helpTip = L('measureTextOutlineTip');

        var randomCheckbox = optionsGroup.add("checkbox", undefined, L('random'));
        randomCheckbox.value = false;

        /* ---- UI: ボタン / Buttons ---- */
        var buttonGroup = dlg.add("group");
        buttonGroup.alignment = "center";
        buttonGroup.alignChildren = ["center", "center"];
        buttonGroup.add("button", undefined, L('cancel'), { name: "cancel" });
        buttonGroup.add("button", undefined, "OK", { name: "ok" });

        /* ---- 内部状態 / Internal state ---- */
        var randomOrderCache = null;
        var outlineMeasureCache = []; /* [{ref:PageItem, b:[l,t,r,b]}] */

        /* レイアウト用に境界を取得（必要なら計測キャッシュ経由） / Get bounds for layout (with optional outline cache) */
        function getBoundsForLayout(item) {
            var dir = getEffectiveDirection();
            if (measureTextOutlineCheckbox.value && dir !== "horizontal") {
                var ob = measureOutlineBoundsOnce(item, outlineMeasureCache);
                if (ob) return ob;
            }
            return getItemBounds(item, usePreviewBoundsCheckbox.value);
        }

        function getHAlignMode() {
            if (rbHNone.value) return "none";
            if (rbHCenter.value) return "center";
            if (rbHRight.value) return "right";
            return "left";
        }

        function getVAlignMode() {
            if (rbVNone.value) return "none";
            if (rbVMiddle.value) return "middle";
            if (rbVBottom.value) return "bottom";
            return "top";
        }

        function getSpacingPt() {
            var v = parseFloat(spacingInput.text);
            if (isNaN(v)) v = 0;
            var unitCode = app.preferences.getIntegerPreference("rulerType");
            return v * getPtFactorFromUnitCode(unitCode);
        }

        /* レイアウト適用 / Apply layout to current selection */
        function applyLayoutToSelection() {
            if (!originalSelection || originalSelection.length === 0) return;

            var direction = getEffectiveDirection();
            var spacingPt = getSpacingPt();

            var order = buildArrangementOrder(originalSelection, direction, randomCheckbox.value, randomOrderCache);
            randomOrderCache = order.cache;
            var items = order.items;
            if (!items.length) return;

            var startBounds = getBoundsForLayout(items[0]);
            var startX = startBounds[0];
            var startY = startBounds[1];

            if (direction === "horizontal") {
                var refHeight = computeMaxHeight(items, getBoundsForLayout);
                applyHorizontalLayout(items, startX, startY, refHeight, spacingPt, getVAlignMode(), getBoundsForLayout);
            } else {
                var refWidth = computeMaxWidth(items, getBoundsForLayout);
                applyVerticalLayout(items, startX, startY, refWidth, spacingPt, getHAlignMode(), getBoundsForLayout);
            }

            if (randomCheckbox.value) {
                applyRandomBaseOffset(items, order.baseLeft, order.baseTop);
            }
        }

        /* ---- 揃えパネル/オプションの有効状態を同期 / Sync enabled state of align panel & options ---- */
        function syncAlignUI() {
            var d = getEffectiveDirection();
            hAlignGroup.visible = true;
            vAlignGroup.visible = true;

            if (d === "horizontal") {
                alignPanel.text = L('alignV');
                hAlignGroup.enabled = false;
                vAlignGroup.enabled = true;
            } else {
                alignPanel.text = L('alignH');
                hAlignGroup.enabled = true;
                vAlignGroup.enabled = false;
            }

            /* テキスト計測は縦並び時かつ選択にテキストがある場合のみ有効 / Enable text-measure only when vertical & selection has text */
            try {
                measureTextOutlineCheckbox.enabled = (d !== "horizontal") && selectionHasText;
            } catch (e) { }

            try { dlg.layout.layout(true); } catch (e) { }
        }

        /* ---- プレビュー更新（前回 preview を app.undo() で巻き戻して再生成） / Preview update via app.undo() ---- */
        var __isPreviewUpdating = false;
        var __lastPreviewAt = 0;
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
        addAlignKeyHandler(dlg, getEffectiveDirection, rbHNone, rbHLeft, rbHCenter, rbHRight, rbVNone, rbVTop, rbVMiddle, rbVBottom, requestPreviewUpdate);
        addAlignKeyHandler(spacingInput, getEffectiveDirection, rbHNone, rbHLeft, rbHCenter, rbHRight, rbVNone, rbVTop, rbVMiddle, rbVBottom, requestPreviewUpdate);

        function onDirChanged() { syncAlignUI(); requestPreviewUpdate(); }
        rbDirAuto.onClick = onDirChanged;
        rbDirV.onClick = onDirChanged;
        rbDirH.onClick = onDirChanged;

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

        var dlgResult = dlg.show();
        try { if (__previewTaskId) app.cancelTask(__previewTaskId); } catch (e) { }
        saveDialogPosition(dlg.location);

        if (dlgResult !== 1) {
            /* キャンセル: プレビューを undo してプリファレンスも戻す / Cancel: undo preview & restore preference */
            cleanupPreview(previewState, activeDocument);
            app.preferences.setBooleanPreference("includeStrokeInBounds", originalIncludeStrokeInBounds);
            app.redraw();
            return false;
        }

        /* 確定: プレビュー分を巻き戻して本実行（undo 履歴を 1 件にまとめる） / OK: undo preview, then commit as single history entry */
        undoPreview(previewState);
        try { app.preferences.setBooleanPreference("includeStrokeInBounds", usePreviewBoundsCheckbox.value); } catch (e) { }
        applyLayoutToSelection();
        try { app.redraw(); } catch (e) { }
        return true;
    }

    // =========================================
    // メイン / Main
    // =========================================

    function main() {
        try {
            var selectedItems = activeDocument.selection;
            if (!selectedItems || selectedItems.length === 0) {
                alert(L('needSelection'));
                return;
            }
            showArrangeDialog();
        } catch (e) {
            alert(L('errorPrefix') + e.message);
        }
    }

    main();

})();