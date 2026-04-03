#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
ArtboardLayerOrganizer.jsx

概要：
ドキュメント内のオブジェクトをアートボード単位で振り分け、
「番号_アートボード名」のレイヤーに整理します。

・各オブジェクトは重心位置を基準に所属アートボードを判定
・ガイドは「_guide」レイヤーへ集約
・どのアートボードにも属さないオブジェクトは「_pasteboard」へ移動
・旧仕様（アートボード名のみのレイヤー）は、新仕様レイヤーへ統合して削除
・レイヤー順はアートボード順（上から1→2→3…）に揃える
・このスクリプトが作成・管理する空レイヤーのみ自動削除
・処理後に空になったサブレイヤーも自動削除
・処理後に空になったトップレベルレイヤーも自動削除
・ダイアログで空レイヤー削除の有無を選択可能
・ロックされたレイヤーを無視するか選択可能
・非表示のレイヤーを無視するか選択可能

更新日：2026-04-04
*/

var SCRIPT_VERSION = "v1.0";

(function () {
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;
    var artboards = doc.artboards;

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    var LABELS = {
        dialogTitle: {
            ja: "アートボードでレイヤー整理",
            en: "Artboard Layer Organizer"
        }
    };

    function L(key) {
        return LABELS[key][lang] || LABELS[key]["en"];
    }

    function showOptionsDialog() {
        var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];
        dlg.margins = [15, 20, 15, 15];

        var pnl = dlg.add("panel", undefined, "オプション");
        pnl.orientation = "column";
        pnl.alignChildren = ["left", "top"];
        pnl.margins = [15, 20, 15, 10];

        var chkRemoveEmpty = pnl.add("checkbox", undefined, "空のレイヤー／サブレイヤーを削除");
        chkRemoveEmpty.value = true;

        var chkIgnoreLocked = pnl.add("checkbox", undefined, "ロックされたレイヤーを無視");
        chkIgnoreLocked.value = true;

        var chkIgnoreHidden = pnl.add("checkbox", undefined, "非表示のレイヤーを無視");
        chkIgnoreHidden.value = true;

        var btns = dlg.add("group");
        btns.orientation = "row";
        btns.alignment = ["fill", "top"];
        btns.add("statictext", undefined, "").alignment = ["fill", "fill"];
        var btnCancel = btns.add("button", undefined, "キャンセル", { name: "cancel" });
        var btnOk = btns.add("button", undefined, "OK", { name: "ok" });
        dlg.defaultElement = btnOk;
        dlg.cancelElement = btnCancel;

        if (dlg.show() !== 1) {
            return null;
        }

        return {
            removeEmptyLayers: chkRemoveEmpty.value,
            ignoreLockedLayers: chkIgnoreLocked.value,
            ignoreHiddenLayers: chkIgnoreHidden.value
        };
    }

    function shouldIgnoreLayer(layer, options) {
        if (!layer || layer.typename !== "Layer") return false;
        if (options.ignoreLockedLayers && layer.locked) return true;
        if (options.ignoreHiddenLayers && !layer.visible) return true;
        return false;
    }

    function hasIgnoredAncestorLayer(obj, options) {
        var parent = obj.parent;
        while (parent) {
            if (parent.typename === "Layer" && shouldIgnoreLayer(parent, options)) {
                return true;
            }
            if (parent.typename === "Document") {
                break;
            }
            parent = parent.parent;
        }
        return false;
    }

    function getOrCreateLayer(name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) {
                return doc.layers[i];
            }
        }
        var layer = doc.layers.add();
        layer.name = name;
        return layer;
    }

    function getArtboardLayerName(index) {
        var abName = artboards[index].name;
        if (!abName || abName === "") {
            abName = "アートボード";
        }
        return (index + 1) + "_" + abName;
    }

    function findLegacyArtboardIndexByLayerName(name) {
        for (var i = 0; i < artboards.length; i++) {
            var abName = artboards[i].name;
            if (!abName || abName === "") {
                abName = "アートボード";
            }
            if (abName === name) {
                return i;
            }
        }
        return -1;
    }

    function moveEntriesToLayerPreservingOrder(entries, targetLayer, processed) {
        var result = {
            moved: 0,
            failed: 0
        };
        var j = entries.length;
        while (j--) {
            try {
                entries[j].item.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
                if (processed) {
                    processed[entries[j].index] = true;
                }
                result.moved++;
            } catch (e) {
                result.failed++;
            }
        }
        return result;
    }

    function collectLayerPageItemsRecursive(layer, outEntries) {
        for (var i = 0; i < layer.layers.length; i++) {
            collectLayerPageItemsRecursive(layer.layers[i], outEntries);
        }
        for (var j = 0; j < layer.pageItems.length; j++) {
            var item = layer.pageItems[j];
            if (item.parent !== layer) continue;
            outEntries.push({ item: item });
        }
    }

    function moveLayerItemsToLayer(sourceLayer, targetLayer) {
        var result = {
            moved: 0,
            failed: 0
        };
        var entries = [];
        collectLayerPageItemsRecursive(sourceLayer, entries);
        var j = entries.length;
        while (j--) {
            try {
                entries[j].item.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
                result.moved++;
            } catch (e) {
                result.failed++;
            }
        }
        return result;
    }

    function removeEmptySubLayers(parentLayer, options) {
        for (var i = parentLayer.layers.length - 1; i >= 0; i--) {
            var sub = parentLayer.layers[i];
            removeEmptySubLayers(sub, options);
            if (shouldIgnoreLayer(sub, options)) continue;
            if (sub.pageItems.length === 0 && sub.layers.length === 0) {
                sub.remove();
            }
        }
    }

    function isInArtboard(item, rect) {
        var b = item.geometricBounds; // [left, top, right, bottom]
        var cx = (b[0] + b[2]) / 2;
        var cy = (b[1] + b[3]) / 2;

        return (
            cx >= rect[0] &&
            cx <= rect[2] &&
            cy <= rect[1] &&
            cy >= rect[3]
        );
    }

    var options = showOptionsDialog();
    if (!options) {
        return;
    }

    // トップレベルの pageItems のみ退避（複合パス等の子を除外）
    var items = [];
    for (var i = 0; i < doc.pageItems.length; i++) {
        var item = doc.pageItems[i];
        var p = item.parent;
        if (p.typename === "Layer" || p.typename === "Document") {
            if (hasIgnoredAncestorLayer(item, options)) continue;
            items.push(item);
        }
    }

    var guideLayer = getOrCreateLayer("_guide");
    var pasteboardLayer = getOrCreateLayer("_pasteboard");
    var managedLayerNames = { "_guide": true, "_pasteboard": true };
    var moved = 0;
    var guideMoved = 0;
    var processed = [];
    var failedMoves = 0;

    for (var a = 0; a < artboards.length; a++) {
        var ab = artboards[a];
        var layerName = getArtboardLayerName(a);
        managedLayerNames[layerName] = true;
        var layer = getOrCreateLayer(layerName);
        var rect = ab.artboardRect;
        var normalEntries = [];
        var guideEntries = [];

        for (var j = 0; j < items.length; j++) {
            if (processed[j]) continue;
            var item = items[j];
            if (hasIgnoredAncestorLayer(item, options)) continue;

            try {
                if (isInArtboard(item, rect)) {
                    if (item.guides) {
                        guideEntries.push({ item: item, index: j });
                    } else {
                        normalEntries.push({ item: item, index: j });
                    }
                }
            } catch (e) {
                // 判定できないものは無視
            }
        }

        var normalResult = moveEntriesToLayerPreservingOrder(normalEntries, layer, processed);
        var guideResult = moveEntriesToLayerPreservingOrder(guideEntries, guideLayer, processed);
        moved += normalResult.moved + guideResult.moved;
        guideMoved += guideResult.moved;
        failedMoves += normalResult.failed + guideResult.failed;
    }

    // アートボード順にレイヤーを上から並べる
    for (var i = artboards.length - 1; i >= 0; i--) {
        var name = getArtboardLayerName(i);
        try {
            var ly = getOrCreateLayer(name);
            ly.zOrder(ZOrderMethod.BRINGTOFRONT);
        } catch (e) {
            // 無視
        }
    }

    // どのアートボードにも属さないアイテムを _pasteboard レイヤーに移動
    var pasteboardNormalEntries = [];
    var pasteboardGuideEntries = [];
    for (var j = 0; j < items.length; j++) {
        if (processed[j]) continue;
        if (hasIgnoredAncestorLayer(items[j], options)) continue;
        if (items[j].guides) {
            pasteboardGuideEntries.push({ item: items[j], index: j });
        } else {
            pasteboardNormalEntries.push({ item: items[j], index: j });
        }
    }

    var pasteboardNormalResult = moveEntriesToLayerPreservingOrder(pasteboardNormalEntries, pasteboardLayer, processed);
    var pasteboardGuideResult = moveEntriesToLayerPreservingOrder(pasteboardGuideEntries, guideLayer, processed);
    moved += pasteboardNormalResult.moved + pasteboardGuideResult.moved;
    guideMoved += pasteboardGuideResult.moved;
    failedMoves += pasteboardNormalResult.failed + pasteboardGuideResult.failed;

    // _guide レイヤーを一番上に移動
    var wasLocked = guideLayer.locked;
    guideLayer.locked = false;
    guideLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
    guideLayer.locked = wasLocked;

    // 旧仕様のアートボード名レイヤーを新仕様レイヤーへ統合
    for (var i = doc.layers.length - 1; i >= 0; i--) {
        var legacyLayer = doc.layers[i];
        if (shouldIgnoreLayer(legacyLayer, options)) continue;
        var legacyArtboardIndex = findLegacyArtboardIndexByLayerName(legacyLayer.name);
        if (legacyArtboardIndex < 0) continue;

        try {
            var targetLayerName = getArtboardLayerName(legacyArtboardIndex);
            managedLayerNames[targetLayerName] = true;
            var targetLayer = getOrCreateLayer(targetLayerName);
            if (legacyLayer !== targetLayer) {
                var legacyMoveResult = moveLayerItemsToLayer(legacyLayer, targetLayer);
                moved += legacyMoveResult.moved;
                failedMoves += legacyMoveResult.failed;
            }
            if (legacyLayer.pageItems.length === 0 && legacyLayer.layers.length === 0 && doc.layers.length > 1) {
                legacyLayer.remove();
            }
        } catch (e) {
            // 統合できないものは無視
        }
    }

    // 旧仕様統合後に、改めてアートボード順へ並べ直す
    for (var i = artboards.length - 1; i >= 0; i--) {
        var orderedName = getArtboardLayerName(i);
        try {
            var orderedLayer = getOrCreateLayer(orderedName);
            orderedLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        } catch (e) {
            // 無視
        }
    }

    // _guide レイヤーを最後に一番上へ戻す
    var guideWasLockedAfterMerge = guideLayer.locked;
    guideLayer.locked = false;
    guideLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
    guideLayer.locked = guideWasLockedAfterMerge;

    var removed = 0;
    if (options.removeEmptyLayers) {
        // 処理後に空になったサブレイヤーを削除
        for (var i = 0; i < doc.layers.length; i++) {
            if (shouldIgnoreLayer(doc.layers[i], options)) continue;
            removeEmptySubLayers(doc.layers[i], options);
        }

        // このスクリプトが管理する空レイヤーのみ削除
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            var ly = doc.layers[i];
            if (shouldIgnoreLayer(ly, options)) continue;
            if (!managedLayerNames[ly.name]) continue;
            removeEmptySubLayers(ly, options);
            if (ly.pageItems.length === 0 && ly.layers.length === 0 && doc.layers.length > 1) {
                ly.remove();
                removed++;
            }
        }
    }

})();