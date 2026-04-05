#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
ArtboardLayerOrganizer.jsx

概要：
ドキュメント内のオブジェクトをアートボード単位で振り分け、
「番号_アートボード名」のレイヤーに整理します。

・各オブジェクトは重心位置を基準に所属アートボードを判定（現在のアートボードのみ対象にすることも可能）
・ガイドは存在する場合のみ「_guide」レイヤーへ集約（必要時にのみ生成）
・どのアートボードにも属さないオブジェクトは「_pasteboard」へ移動（必要時のみ生成）
・旧仕様（アートボード名のみのレイヤー）は、新仕様レイヤーへ統合して削除
・旧仕様レイヤー統合後も、レイヤー順はアートボード順（上から1→2→3…）に揃える
・ダイアログで対象範囲（現在のアートボード／すべてのアートボード）を選択可能
・ダイアログでレイヤー名に含める要素（アートボード番号／アートボード名）と区切り文字を選択可能
・ダイアログでロック／非表示のレイヤー・オブジェクトを対象外にするか選択可能
・ダイアログで空のレイヤー／サブレイヤー削除の有無を選択可能
・処理後に空になったトップレベルレイヤーおよびサブレイヤーを削除

更新日：2026-04-06
*/

var SCRIPT_VERSION = "v1.2";

(function () {
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    var LABELS = {
        dialog: {
            title: {
                ja: "アートボードでレイヤー整理",
                en: "Artboard Layer Organizer"
            }
        },
        checkbox: {
            removeEmpty: {
                ja: "空のレイヤー{slash}サブレイヤーを削除",
                en: "Remove empty layers{slash}sub-layers"
            },
            includeArtboardNumber: {
                ja: "アートボード番号",
                en: "Artboard Number"
            },
            includeArtboardName: {
                ja: "アートボード名",
                en: "Artboard Name"
            },
            ignoreLocked: {
                ja: "ロックされたレイヤー",
                en: "Locked layers"
            },
            ignoreLockedObjects: {
                ja: "ロックされたオブジェクト",
                en: "Locked objects"
            },
            ignoreHidden: {
                ja: "非表示のレイヤー",
                en: "Hidden layers"
            },
            ignoreHiddenObjects: {
                ja: "非表示のオブジェクト",
                en: "Hidden objects"
            }
        },
        panel: {
            exclude: {
                ja: "対象外にする",
                en: "Exclude"
            },
            target: {
                ja: "対象",
                en: "Target"
            },
            layerName: {
                ja: "レイヤー名",
                en: "Layer Name"
            }
        },
        dropdown: {
            separatorUnderscore: {
                ja: "アンダースコア (_) ",
                en: "Underscore (_)"
            },
            separatorHyphen: {
                ja: "ハイフン (-)",
                en: "Hyphen (-)"
            },
            separatorSpace: {
                ja: "半角スペース",
                en: "Space"
            },
            separatorNone: {
                ja: "なし",
                en: "None"
            }
        },
        fallback: {
            artboard: {
                ja: "アートボード",
                en: "Artboard"
            }
        },
        radio: {
            currentArtboardOnly: {
                ja: "現在のアートボードのみ",
                en: "Current artboard"
            },
            allArtboards: {
                ja: "すべてのアートボード",
                en: "All artboards"
            }
        },
        button: {
            cancel: {
                ja: "キャンセル",
                en: "Cancel"
            },
            ok: {
                ja: "OK",
                en: "OK"
            }
        },
        alert: {
            noDoc: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            failed: {
                ja: "一部のオブジェクトを移動できませんでした。\n移動失敗: ",
                en: "Some objects could not be moved.\nFailed moves: "
            }
        }
    };

    function L(path) {
        var parts = path.split(".");
        var obj = LABELS;
        for (var i = 0; i < parts.length; i++) {
            obj = obj[parts[i]];
            if (!obj) return path;
        }
        var text = obj[lang] || obj["en"];
        if (!text) return path;
        return applyUISymbols(text);
    }

    function applyUISymbols(text) {
        return text
            .replace(/\{slash\}/g, uiSymbol("slash"))
            .replace(/\{colon\}/g, uiSymbol("colon"))
            .replace(/\{comma\}/g, uiSymbol("comma"))
            .replace(/\{openParen\}/g, uiSymbol("openParen"))
            .replace(/\{closeParen\}/g, uiSymbol("closeParen"));
    }

    function uiSymbol(name) {
        if (lang === "ja") {
            switch (name) {
                case "slash": return "／";
                case "colon": return "：";
                case "comma": return "、";
                case "openParen": return "（";
                case "closeParen": return "）";
            }
        }
        switch (name) {
            case "slash": return "/";
            case "colon": return ":";
            case "comma": return ", ";
            case "openParen": return "(";
            case "closeParen": return ")";
        }
        return "";
    }

    if (app.documents.length === 0) {
        alert(L("alert.noDoc"));
        return;
    }

    var doc = app.activeDocument;
    var artboards = doc.artboards;

    function showOptionsDialog() {
        var dlg = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];
        dlg.margins = [15, 20, 15, 15];

        var pnlTarget = dlg.add("panel", undefined, L("panel.target"));
        pnlTarget.orientation = "column";
        pnlTarget.alignChildren = ["left", "top"];
        pnlTarget.margins = [15, 20, 15, 10];

        var rbCurrentArtboardOnly = pnlTarget.add("radiobutton", undefined, L("radio.currentArtboardOnly"));
        var rbAllArtboards = pnlTarget.add("radiobutton", undefined, L("radio.allArtboards"));

        if (artboards.length <= 1) {
            rbCurrentArtboardOnly.value = true;
            rbAllArtboards.enabled = false;
        } else {
            rbAllArtboards.value = true;
        }

        var pnlLayerName = dlg.add("panel", undefined, L("panel.layerName"));
        pnlLayerName.orientation = "column";
        pnlLayerName.alignChildren = ["left", "top"];
        pnlLayerName.margins = [15, 20, 15, 10];

        var chkIncludeArtboardNumber = pnlLayerName.add("checkbox", undefined, L("checkbox.includeArtboardNumber"));
        chkIncludeArtboardNumber.value = true;

        var grpSeparator = pnlLayerName.add("group");
        grpSeparator.orientation = "row";
        grpSeparator.alignChildren = ["left", "center"];
        var ddSeparator = grpSeparator.add("dropdownlist", undefined, [
            L("dropdown.separatorUnderscore"),
            L("dropdown.separatorHyphen"),
            L("dropdown.separatorSpace"),
            L("dropdown.separatorNone")
        ]);
        ddSeparator.selection = 0;

        function updateSeparatorEnabled() {
            grpSeparator.enabled = chkIncludeArtboardNumber.value;
        }

        updateSeparatorEnabled();

        var chkIncludeArtboardName = pnlLayerName.add("checkbox", undefined, L("checkbox.includeArtboardName"));
        chkIncludeArtboardName.value = true;

        var pnl = dlg.add("panel", undefined, L("panel.exclude"));
        pnl.orientation = "column";
        pnl.alignChildren = ["left", "top"];
        pnl.margins = [15, 20, 15, 10];

        var chkIgnoreLocked = pnl.add("checkbox", undefined, L("checkbox.ignoreLocked"));
        chkIgnoreLocked.value = true;

        var chkIgnoreLockedObjects = pnl.add("checkbox", undefined, L("checkbox.ignoreLockedObjects"));
        chkIgnoreLockedObjects.value = true;

        var chkIgnoreHidden = pnl.add("checkbox", undefined, L("checkbox.ignoreHidden"));
        chkIgnoreHidden.value = true;

        var chkIgnoreHiddenObjects = pnl.add("checkbox", undefined, L("checkbox.ignoreHiddenObjects"));
        chkIgnoreHiddenObjects.value = true;

        var grpRemoveEmpty = dlg.add("group");
        grpRemoveEmpty.orientation = "row";
        grpRemoveEmpty.alignment = ["center", "top"];
        var chkRemoveEmpty = grpRemoveEmpty.add("checkbox", undefined, L("checkbox.removeEmpty"));
        chkRemoveEmpty.value = true;

        var btns = dlg.add("group");
        btns.orientation = "row";
        btns.alignment = ["center", "top"];
        var btnCancel = btns.add("button", undefined, L("button.cancel"), { name: "cancel" });
        var btnOk = btns.add("button", undefined, L("button.ok"), { name: "ok" });
        dlg.defaultElement = btnOk;
        dlg.cancelElement = btnCancel;

        chkIncludeArtboardNumber.onClick = updateSeparatorEnabled;

        if (dlg.show() !== 1) {
            return null;
        }

        return {
            removeEmptyLayers: chkRemoveEmpty.value,
            currentArtboardOnly: rbCurrentArtboardOnly.value,
            includeArtboardNumber: chkIncludeArtboardNumber.value,
            includeArtboardName: chkIncludeArtboardName.value,
            layerNameSeparatorIndex: ddSeparator.selection ? ddSeparator.selection.index : 0,
            ignoreLockedLayers: chkIgnoreLocked.value,
            ignoreLockedObjects: chkIgnoreLockedObjects.value,
            ignoreHiddenLayers: chkIgnoreHidden.value,
            ignoreHiddenObjects: chkIgnoreHiddenObjects.value
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

    // PageItem の非表示判定は visible ではなく hidden を使う
    // 一部の PageItem は locked / hidden 参照で例外になることがあるため、各プロパティごとに個別に判定する
    function shouldIgnoreObject(item, options) {
        if (options.ignoreLockedObjects) {
            try {
                if (item.locked) return true;
            } catch (e) {
                // locked 判定不可のものは locked 条件では除外しない
            }
        }
        if (options.ignoreHiddenObjects) {
            try {
                if (item.hidden) return true;
            } catch (e) {
                // hidden 判定不可のものは hidden 条件では除外しない
            }
        }
        return false;
    }

    function findLayerByName(name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) {
                return doc.layers[i];
            }
        }
        return null;
    }

    function getOrCreateLayer(name) {
        var layer = findLayerByName(name);
        if (layer) {
            return layer;
        }
        layer = doc.layers.add();
        layer.name = name;
        return layer;
    }

    function getOrCreateGuideLayer() {
        if (!guideLayer) {
            guideLayer = getOrCreateLayer("_guide");
        }
        return guideLayer;
    }

    function getLayerNameSeparator() {
        var index = options && typeof options.layerNameSeparatorIndex === "number" ? options.layerNameSeparatorIndex : 0;
        switch (index) {
            case 1: return "-";
            case 2: return " ";
            case 3: return "";
            default: return "_";
        }
    }

    function getArtboardLayerName(index) {
        var parts = [];
        var abName = artboards[index].name;
        if (!abName || abName === "") {
            abName = L("fallback.artboard");
        }

        if (!options || options.includeArtboardNumber !== false) {
            parts.push(String(index + 1));
        }
        if (!options || options.includeArtboardName !== false) {
            parts.push(abName);
        }

        if (parts.length === 0) {
            parts.push(String(index + 1));
            parts.push(abName);
        }
        return parts.join(getLayerNameSeparator());
    }

    function findLegacyArtboardIndexByLayerName(name) {
        for (var i = 0; i < artboards.length; i++) {
            var abName = artboards[i].name;
            if (!abName || abName === "") {
                abName = L("fallback.artboard");
            }
            if (abName === name) {
                return i;
            }
        }
        return -1;
    }

    function moveCollectedEntries(entries, targetLayer, processed) {
        var result = {
            moved: 0,
            failed: 0
        };
        var j = entries.length;
        while (j--) {
            try {
                entries[j].item.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
                if (processed && typeof entries[j].index === "number") {
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
        var entries = [];
        collectLayerPageItemsRecursive(sourceLayer, entries);
        return moveCollectedEntries(entries, targetLayer, null);
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
            if (shouldIgnoreObject(item, options)) continue;
            items.push(item);
        }
    }

    // _guide は既存レイヤーを先に拾い、必要になった時点で getOrCreateGuideLayer() が参照を更新する
    var guideLayer = findLayerByName("_guide");
    var pasteboardLayer = null;
    var processed = [];
    var failedMoves = 0;

    var startIndex = 0;
    var endIndex = artboards.length;
    if (options.currentArtboardOnly) {
        startIndex = doc.artboards.getActiveArtboardIndex();
        endIndex = startIndex + 1;
    }

    for (var a = startIndex; a < endIndex; a++) {
        var ab = artboards[a];
        var layerName = getArtboardLayerName(a);
        var layer = getOrCreateLayer(layerName);
        var rect = ab.artboardRect;
        var normalEntries = [];
        var guideEntries = [];

        for (var j = 0; j < items.length; j++) {
            if (processed[j]) continue;
            var item = items[j];

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

        var normalResult = moveCollectedEntries(normalEntries, layer, processed);
        var guideResult = { moved: 0, failed: 0 };
        if (guideEntries.length > 0) {
            guideResult = moveCollectedEntries(guideEntries, getOrCreateGuideLayer(), processed);
        }
        failedMoves += normalResult.failed + guideResult.failed;
    }

    // 対象アートボードのレイヤーを上へ並べる
    for (var i = endIndex - 1; i >= startIndex; i--) {
        var name = getArtboardLayerName(i);
        try {
            var ly = getOrCreateLayer(name);
            ly.zOrder(ZOrderMethod.BRINGTOFRONT);
        } catch (e) {
            // 無視
        }
    }

    // 全アートボード処理時のみ、どのアートボードにも属さないアイテムを _pasteboard レイヤーに移動
    if (!options.currentArtboardOnly) {
        var pasteboardNormalEntries = [];
        var pasteboardGuideEntries = [];
        for (var j = 0; j < items.length; j++) {
            if (processed[j]) continue;
            if (items[j].guides) {
                pasteboardGuideEntries.push({ item: items[j], index: j });
            } else {
                pasteboardNormalEntries.push({ item: items[j], index: j });
            }
        }

        var pasteboardNormalResult = { moved: 0, failed: 0 };
        if (pasteboardNormalEntries.length > 0) {
            pasteboardLayer = getOrCreateLayer("_pasteboard");
            pasteboardNormalResult = moveCollectedEntries(pasteboardNormalEntries, pasteboardLayer, processed);
        }
        var pasteboardGuideResult = { moved: 0, failed: 0 };
        if (pasteboardGuideEntries.length > 0) {
            pasteboardGuideResult = moveCollectedEntries(pasteboardGuideEntries, getOrCreateGuideLayer(), processed);
        }
        failedMoves += pasteboardNormalResult.failed + pasteboardGuideResult.failed;
    }

    // _guide レイヤーを一番上に移動
    if (guideLayer) {
        var wasLocked = guideLayer.locked;
        guideLayer.locked = false;
        guideLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        guideLayer.locked = wasLocked;
    }

    // 旧仕様のアートボード名レイヤーを新仕様レイヤーへ統合
    if (!options.currentArtboardOnly) {
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            var legacyLayer = doc.layers[i];
            if (shouldIgnoreLayer(legacyLayer, options)) continue;
            var legacyArtboardIndex = findLegacyArtboardIndexByLayerName(legacyLayer.name);
            if (legacyArtboardIndex < 0) continue;

            try {
                var targetLayerName = getArtboardLayerName(legacyArtboardIndex);
                var targetLayer = getOrCreateLayer(targetLayerName);
                if (legacyLayer !== targetLayer) {
                    var legacyMoveResult = moveLayerItemsToLayer(legacyLayer, targetLayer);
                    failedMoves += legacyMoveResult.failed;
                }
                if (legacyLayer.pageItems.length === 0 && legacyLayer.layers.length === 0 && doc.layers.length > 1) {
                    legacyLayer.remove();
                }
            } catch (e) {
                // 統合できないものは無視
            }
        }
    }

    // 旧仕様統合後に、対象アートボードのレイヤーを改めて上へ並べ直す
    for (var i = endIndex - 1; i >= startIndex; i--) {
        var orderedName = getArtboardLayerName(i);
        try {
            var orderedLayer = getOrCreateLayer(orderedName);
            orderedLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        } catch (e) {
            // 無視
        }
    }

    // _guide レイヤーを最後に一番上へ戻す
    if (guideLayer) {
        var guideWasLockedAfterMerge = guideLayer.locked;
        guideLayer.locked = false;
        guideLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        guideLayer.locked = guideWasLockedAfterMerge;
    }

    if (options.removeEmptyLayers) {
        // 処理後に空になったサブレイヤーを削除
        for (var i = 0; i < doc.layers.length; i++) {
            if (shouldIgnoreLayer(doc.layers[i], options)) continue;
            removeEmptySubLayers(doc.layers[i], options);
        }

        // 処理後に空のトップレベルレイヤーを削除
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            var ly = doc.layers[i];
            if (shouldIgnoreLayer(ly, options)) continue;
            removeEmptySubLayers(ly, options);
            if (ly.pageItems.length === 0 && ly.layers.length === 0 && doc.layers.length > 1) {
                ly.remove();
            }
        }
    }

    if (failedMoves > 0) {
        alert(L("alert.failed") + failedMoves);
    }

})();