#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

var SCRIPT_VERSION = "v1.0.0";

(function () {
    var TEXT_LAYER_NAME = "text_all";

    var lang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* 日英ラベル定義 / Define Japanese-English labels */
    var LABELS = {
        dialogTitle: {
            ja: "Excelデータを整形",
            en: "Format Excel Data"
        },
        releaseMask: { ja: "マスク解除", en: "Release mask" },
        optionsPanelTitle: { ja: "オプション", en: "Options" },
        textPanelTitle: { ja: "テキスト", en: "Text" },
        cellBgPanelTitle: { ja: "セル背景", en: "Cell Background" },
        moveTextToLayer: {
            ja: "テキストを別レイヤーに",
            en: "Move text to separate layer"
        },
        setTextK100: {
            ja: "印刷用の「黒」（K100）に",
            en: "Set text to print black (K100)"
        },
        removeDuplicateTexts: {
            ja: "重複テキストを削除",
            en: "Remove duplicate texts"
        },
        removeSmallObjects: {
            ja: "小さいオブジェクトを削除",
            en: "Remove small objects"
        },
        adjustCellBackground: {
            ja: "セル背景を調整",
            en: "Adjust cell background"
        },
        equalizeHeights: {
            ja: "高さを均等に",
            en: "Equalize heights"
        },
        ok: { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects selected." }
    };

    function L(key) {
        return LABELS[key][lang];
    }

    function main() {
        var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = "fill";
        dlg.spacing = 10;
        dlg.margins = 20;

        var optionsPanel = dlg.add("panel", undefined, L('optionsPanelTitle'));
        optionsPanel.orientation = "column";
        optionsPanel.alignChildren = "left";
        optionsPanel.margins = [15, 20, 15, 15];

        var releaseMaskCheckbox = optionsPanel.add("checkbox", undefined, L('releaseMask'));
        releaseMaskCheckbox.value = true;

        var removeSmallCheckbox = optionsPanel.add("checkbox", undefined, L('removeSmallObjects'));
        removeSmallCheckbox.value = true;

        var textPanel = dlg.add("panel", undefined, L('textPanelTitle'));
        textPanel.orientation = "column";
        textPanel.alignChildren = "left";
        textPanel.margins = [15, 20, 15, 15];

        var moveTextCheckbox = textPanel.add("checkbox", undefined, L('moveTextToLayer'));
        moveTextCheckbox.value = true;

        var setK100Checkbox = textPanel.add("checkbox", undefined, L('setTextK100'));
        setK100Checkbox.value = true;

        var removeDuplicateCheckbox = textPanel.add("checkbox", undefined, L('removeDuplicateTexts'));
        removeDuplicateCheckbox.value = true;

        var cellBgPanel = dlg.add("panel", undefined, L('cellBgPanelTitle'));
        cellBgPanel.orientation = "column";
        cellBgPanel.alignChildren = "left";
        cellBgPanel.margins = [15, 20, 15, 15];

        var adjustCellBgCheckbox = cellBgPanel.add("checkbox", undefined, L('adjustCellBackground'));
        adjustCellBgCheckbox.value = true;

        var equalizeHeightsCheckbox = cellBgPanel.add("checkbox", undefined, L('equalizeHeights'));
        equalizeHeightsCheckbox.value = true;

        adjustCellBgCheckbox.onClick = function () {
            equalizeHeightsCheckbox.enabled = adjustCellBgCheckbox.value;
        };

        var btnGroup = dlg.add("group");
        btnGroup.orientation = "row";
        btnGroup.alignment = "center";
        btnGroup.margins = [0, 10, 0, 0];

        var btnCancel = btnGroup.add("button", undefined, L('cancel'), { name: "cancel" });
        var btnOk = btnGroup.add("button", undefined, L('ok'), { name: "ok" });

        btnOk.onClick = function () {
            var releaseMask = releaseMaskCheckbox.value;
            var moveText = moveTextCheckbox.value;
            var setK100 = setK100Checkbox.value;
            var removeDuplicate = removeDuplicateCheckbox.value;
            var removeSmall = removeSmallCheckbox.value;
            var adjustCellBg = adjustCellBgCheckbox.value;
            var equalizeHeights = equalizeHeightsCheckbox.value;
            dlg.close();
            executeRelease(releaseMask, moveText, setK100, removeDuplicate, removeSmall, adjustCellBg, equalizeHeights);
        };

        dlg.show();
    }

    function executeRelease(releaseMask, moveText, setK100, removeDuplicate, removeSmall, adjustCellBg, equalizeHeights) {
        if (!app.documents.length || !app.activeDocument.selection.length) {
            alert(L('noSelection'));
            return;
        }

        var doc = app.activeDocument;
        var selectionItems = doc.selection;

        if (removeDuplicate) {
            processClipGroupTexts(selectionItems);
        }

        if (releaseMask) {
            for (var i = 0; i < selectionItems.length; i++) {
                var currentItem = selectionItems[i];
                if (currentItem.typename === "GroupItem" && currentItem.clipped === true) {
                    for (var j = 0; j < currentItem.pageItems.length; j++) {
                        var item = currentItem.pageItems[j];
                        if (item.typename === "PathItem" && item.clipping) {
                            item.remove();
                            break;
                        }
                    }
                    ungroup(currentItem);
                }
            }
        }

        if (removeSmall) {
            removeSmallPathItems(doc);
        }

        if (moveText) {
            moveTextsToLayer(doc, TEXT_LAYER_NAME);
        }

        if (setK100) {
            applyK100ToAllTexts(doc);
        }

        if (adjustCellBg) {
            autoSelectAndMerge({
                layerName: "_path",
                excludeLayerNames: [TEXT_LAYER_NAME],
                silent: true
            });
            if (equalizeHeights) {
                equalizeRectangleHeightsInLayer(doc, "_path");
            }
        }

        hideLayerByName(doc, TEXT_LAYER_NAME);
        hideLayerByName(doc, "_path");
    }

    function hideLayerByName(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) {
                try { doc.layers[i].visible = false; } catch (e) { }
                return;
            }
        }
    }

    function removeSmallPathItems(doc) {
        var minTextSize = null;
        for (var i = 0; i < doc.textFrames.length; i++) {
            var tf = doc.textFrames[i];
            try {
                var s = tf.textRange.characterAttributes.size;
                if (typeof s === 'number' && s > 0) {
                    if (minTextSize === null || s < minTextSize) minTextSize = s;
                }
            } catch (e) { }
        }
        if (minTextSize === null || minTextSize <= 0) return;

        var threshold = minTextSize / 2;
        var toRemove = [];
        for (var j = 0; j < doc.pathItems.length; j++) {
            var p = doc.pathItems[j];
            try {
                if (p.clipping) continue;
                if (p.locked || (p.layer && p.layer.locked)) continue;
                var pb = p.geometricBounds;
                var pw = pb[2] - pb[0];
                var ph = pb[1] - pb[3];
                if (pw < threshold && ph < threshold) {
                    toRemove.push(p);
                }
            } catch (e) { }
        }
        for (var k = 0; k < toRemove.length; k++) {
            try { toRemove[k].remove(); } catch (e) { }
        }
    }

    function processClipGroupTexts(selectionItems) {
        for (var i = 0; i < selectionItems.length; i++) {
            var item = selectionItems[i];
            if (item.typename !== "GroupItem" || item.clipped !== true) continue;
            deduplicateTextsInGroup(item);
            concatenateTextsInGroup(item);
        }
    }

    function deduplicateTextsInGroup(group) {
        var texts = [];
        collectTextFrames(group, texts);
        if (texts.length < 2) return;
        var seen = {};
        var toRemove = [];
        for (var k = texts.length - 1; k >= 0; k--) {
            var key = "k:" + texts[k].contents;
            if (seen[key]) {
                toRemove.push(texts[k]);
            } else {
                seen[key] = true;
            }
        }
        for (var r = 0; r < toRemove.length; r++) {
            try { toRemove[r].remove(); } catch (e) { }
        }
    }

    function concatenateTextsInGroup(group) {
        var texts = [];
        collectTextFrames(group, texts);
        if (texts.length < 2) return;
        texts.sort(function (a, b) {
            var ay, by, ax, bx;
            try {
                ay = a.position[1]; by = b.position[1];
                ax = a.position[0]; bx = b.position[0];
            } catch (e) { return 0; }
            if (Math.abs(ay - by) > 0.5) return by - ay;
            return ax - bx;
        });
        var combined = "";
        for (var i = 0; i < texts.length; i++) {
            combined += texts[i].contents;
        }
        try {
            texts[0].contents = combined;
        } catch (e) { return; }
        for (var j = 1; j < texts.length; j++) {
            try { texts[j].remove(); } catch (e) { }
        }
    }

    function collectTextFrames(group, out) {
        for (var i = 0; i < group.pageItems.length; i++) {
            var p = group.pageItems[i];
            try {
                if (p.typename === "TextFrame") {
                    out.push(p);
                } else if (p.typename === "GroupItem") {
                    collectTextFrames(p, out);
                }
            } catch (e) { }
        }
    }

    function applyK100ToAllTexts(doc) {
        var black = getK100Black();
        for (var i = 0; i < doc.textFrames.length; i++) {
            var tf = doc.textFrames[i];
            try {
                tf.textRange.characterAttributes.fillColor = black;
            } catch (e) { }
        }
    }

    function moveTextsToLayer(doc, layerName) {
        var targetLayer = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) {
                targetLayer = doc.layers[i];
                break;
            }
        }
        if (!targetLayer) {
            targetLayer = doc.layers.add();
            targetLayer.name = layerName;
        }
        if (targetLayer.locked) targetLayer.locked = false;
        if (!targetLayer.visible) targetLayer.visible = true;

        var texts = [];
        for (var j = 0; j < doc.textFrames.length; j++) {
            texts.push(doc.textFrames[j]);
        }
        for (var k = 0; k < texts.length; k++) {
            var t = texts[k];
            if (t.parent === targetLayer) continue;
            t.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
        }
    }

    function ungroup(groupItem) {
        var parent = groupItem.parent;
        while (groupItem.pageItems.length > 0) {
            groupItem.pageItems[0].move(parent, ElementPlacement.PLACEATEND);
        }
        groupItem.remove();
    }

    function getK100Black() {
        var cmykColor = new CMYKColor();
        cmykColor.cyan = 0;
        cmykColor.magenta = 0;
        cmykColor.yellow = 0;
        cmykColor.black = 100;
        return cmykColor;
    }

    function getModeTextSize(doc) {
        var counts = {};
        for (var t = 0; t < doc.textFrames.length; t++) {
            try {
                var chars = doc.textFrames[t].characters;
                for (var c = 0; c < chars.length; c++) {
                    try {
                        var key = Math.round(chars[c].size * 10) / 10;
                        counts[key] = (counts[key] || 0) + 1;
                    } catch (e) { }
                }
            } catch (e) { }
        }
        var modeSize = 0, modeCount = 0;
        for (var k in counts) {
            if (counts.hasOwnProperty(k) && counts[k] > modeCount) {
                modeCount = counts[k];
                modeSize = parseFloat(k);
            }
        }
        return modeSize;
    }

    function isRectangleLike(item) {
        if (!item || item.typename !== "PathItem") return false;
        if (!item.closed) return false;
        if (item.pathPoints.length !== 4) return false;
        return true;
    }

    function containsItem(arr, item) {
        for (var i = 0; i < arr.length; i++) {
            if (arr[i] === item) return true;
        }
        return false;
    }

    function collectRectangles(doc, opts) {
        opts = opts || {};
        var excludeNames = opts.excludeLayerNames || [];
        var shortSideMin = (typeof opts.shortSideMin === "number") ? opts.shortSideMin : 0;
        var customFilter = opts.filter || null;
        var result = [];

        function walk(items) {
            for (var i = 0; i < items.length; i++) {
                var it = items[i];
                try {
                    if (it.typename === "GroupItem") {
                        walk(it.pageItems);
                    } else if (isRectangleLike(it)) {
                        var w = it.width, h = it.height;
                        var shortSide = (w < h) ? w : h;
                        if (shortSide > shortSideMin) {
                            if (customFilter === null || customFilter(it)) {
                                result.push(it);
                            }
                        }
                    }
                } catch (e) { }
            }
        }

        for (var li = 0; li < doc.layers.length; li++) {
            var ly = doc.layers[li];
            var skip = false;
            for (var ei = 0; ei < excludeNames.length; ei++) {
                if (ly.name === excludeNames[ei]) { skip = true; break; }
            }
            if (skip || ly.locked || !ly.visible) continue;
            walk(ly.pageItems);
        }
        return result;
    }

    function ensureBackLayer(doc, layerName) {
        var layer = null;
        for (var l = 0; l < doc.layers.length; l++) {
            if (doc.layers[l].name === layerName) { layer = doc.layers[l]; break; }
        }
        if (layer === null) {
            layer = doc.layers.add();
            layer.name = layerName;
        }
        try { layer.zOrder(ZOrderMethod.SENDTOBACK); } catch (e) { }
        return layer;
    }

    function expandByAppearance(doc, items) {
        var found = [];
        for (var n = 0; n < items.length; n++) {
            try {
                doc.selection = null;
                items[n].selected = true;
                app.executeMenuCommand('Find Appearance menu item');
                for (var j = 0; j < doc.selection.length; j++) {
                    if (!containsItem(found, doc.selection[j])) {
                        found.push(doc.selection[j]);
                    }
                }
            } catch (e) { }
        }
        return found;
    }

    function moveItemsToLayer(items, layer) {
        for (var k = items.length - 1; k >= 0; k--) {
            try { items[k].move(layer, ElementPlacement.PLACEATBEGINNING); } catch (e) { }
        }
    }

    function setSelection(doc, items) {
        doc.selection = null;
        for (var m = 0; m < items.length; m++) {
            try { items[m].selected = true; } catch (e) { }
        }
    }

    function collectRectanglesInLayer(doc, layerName) {
        var result = [];
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name !== layerName) continue;
            var ly = doc.layers[i];
            if (ly.locked || !ly.visible) return result;
            walkRectangles(ly.pageItems, result);
            break;
        }
        return result;
    }

    function walkRectangles(items, out) {
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            try {
                if (it.typename === "GroupItem") {
                    walkRectangles(it.pageItems, out);
                } else if (isRectangleLike(it)) {
                    out.push(it);
                }
            } catch (e) { }
        }
    }

    function equalizeRectangleHeightsInLayer(doc, layerName) {
        var rects = collectRectanglesInLayer(doc, layerName);
        if (rects.length < 1) return;

        var minX = null, maxX = null;
        for (var i = 0; i < rects.length; i++) {
            try {
                var gb = rects[i].geometricBounds;
                if (minX === null || gb[0] < minX) minX = gb[0];
                if (maxX === null || gb[2] > maxX) maxX = gb[2];
            } catch (e) { }
        }
        if (minX === null || maxX === null) return;

        var newHeight = (maxX - minX) / rects.length;
        if (newHeight <= 0) return;

        for (var j = 0; j < rects.length; j++) {
            try {
                var rect = rects[j];
                if (Math.abs(rect.height - newHeight) < 0.01) continue;
                var pos = rect.position;
                rect.height = newHeight;
                rect.position = pos;
            } catch (e) { }
        }

        stackRectanglesVertically(rects);
    }

    function stackRectanglesVertically(rects) {
        if (rects.length < 2) return;

        var tolerance = 0.5;
        var columns = [];

        for (var i = 0; i < rects.length; i++) {
            var rect = rects[i];
            var left;
            try { left = rect.position[0]; } catch (e) { continue; }
            var placed = false;
            for (var c = 0; c < columns.length; c++) {
                if (Math.abs(columns[c].left - left) < tolerance) {
                    columns[c].items.push(rect);
                    placed = true;
                    break;
                }
            }
            if (!placed) {
                columns.push({ left: left, items: [rect] });
            }
        }

        for (var k = 0; k < columns.length; k++) {
            var col = columns[k];
            if (col.items.length < 2) continue;
            col.items.sort(function (a, b) {
                try { return b.position[1] - a.position[1]; } catch (e) { return 0; }
            });
            for (var j = 1; j < col.items.length; j++) {
                try {
                    var prev = col.items[j - 1];
                    var curr = col.items[j];
                    var nextTop = prev.position[1] - prev.height;
                    curr.position = [curr.position[0], nextTop];
                } catch (e) { }
            }
        }
    }

    function autoSelectAndMerge(opts) {
        opts = opts || {};
        if (app.documents.length === 0) return null;
        var doc = app.activeDocument;
        var layerName = opts.layerName || "_path";

        var threshold;
        if (typeof opts.shortSideMin === "function") {
            threshold = opts.shortSideMin(doc);
        } else if (typeof opts.shortSideMin === "number") {
            threshold = opts.shortSideMin;
        } else {
            threshold = getModeTextSize(doc);
        }
        if (!threshold || threshold <= 0) return null;

        var excludeNames = [layerName];
        if (opts.excludeLayerNames) {
            for (var ex = 0; ex < opts.excludeLayerNames.length; ex++) {
                excludeNames.push(opts.excludeLayerNames[ex]);
            }
        }

        var rectangles = collectRectangles(doc, {
            excludeLayerNames: excludeNames,
            shortSideMin: threshold
        });
        if (rectangles.length === 0) return null;

        setSelection(doc, rectangles);

        var targets;
        if (opts.runFindAppearance === false) {
            targets = rectangles;
        } else {
            targets = expandByAppearance(doc, rectangles);
            if (targets.length === 0) return null;
        }

        var targetLayer = ensureBackLayer(doc, layerName);
        moveItemsToLayer(targets, targetLayer);
        setSelection(doc, targets);

        if (opts.runGroup !== false) { try { app.executeMenuCommand('group'); } catch (e) { } }
        if (opts.runMerge !== false) { try { app.executeMenuCommand('Live Pathfinder Merge'); } catch (e) { } }
        if (opts.runExpand !== false) { try { app.executeMenuCommand('expandStyle'); } catch (e) { } }

        return targets;
    }

    main();
})();
