#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
ArtboardLayerOrganizer.jsx

概要：
ドキュメント内のオブジェクトをアートボード単位で振り分け、
「番号_アートボード名」のレイヤーに整理します。

[振り分け]
- 各オブジェクトは重心位置を基準に所属アートボードを判定
- ガイドは「_guide」、どのアートボードにも属さないオブジェクトは「_pasteboard」へ集約（必要時のみ生成）
- ガイドは「指定レイヤー」除外を無視して _guide へ集約（ロック／非表示の祖先は尊重）

[旧仕様レイヤーの整理]
- アートボード名のみのレイヤーは新仕様レイヤーへ統合して削除
- 統合後もレイヤー順はアートボード順（上から 1→2→3…）に揃える

[ダイアログで選択]
- 対象範囲（現在のアートボード／すべてのアートボード）
- レイヤー名に含める要素（アートボード番号／アートボード名）と区切り文字
- ロック／非表示のレイヤー・オブジェクトを対象外にするか
- 「指定レイヤー」欄にレイヤー名（,または、区切り）を入れて整理対象外に
- 空のレイヤー／サブレイヤー削除の有無

[保護・自動処理]
- 「_guide」「_pasteboard」は空でも削除しないシステム管理レイヤー
- 移動先レイヤーがロック／非表示でも自動で解除して移動し、元の状態に戻す
- 処理後に空のトップレベルレイヤー・サブレイヤーを削除（保護レイヤーを除く）

更新日：2026-05-26
*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.3.0";

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* システム管理レイヤー名（空でも削除しない） / System-managed layer names (never deleted) */
    var GUIDE_LAYER_NAME = "_guide";
    var PASTEBOARD_LAYER_NAME = "_pasteboard";

    /* パネル共通レイアウト / Common panel layout */
    var PANEL_MARGINS = [15, 20, 15, 10];
    var PANEL_SPACING = 8;

    /* ダイアログの初期値。必要に応じて編集 / Dialog defaults; edit as needed */
    var DEFAULTS = {
        removeEmptyLayers: true,
        includeArtboardNumber: true,
        includeArtboardName: true,
        useSeparator: true,
        separatorIndex: 0,      // 0:_  1:-  2:スペース  3:なし
        ignoreLockedLayers: true,
        ignoreLockedObjects: true,
        ignoreHiddenLayers: true,
        ignoreHiddenObjects: true,
        excludedLayerNames: "bg"    // , または 、 区切り / Comma-separated
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* UI 言語を判定 / Detect UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "アートボードでレイヤー整理", en: "Artboard Layer Organizer" }
        },
        checkbox: {
            removeEmpty: { ja: "空のレイヤー{slash}サブレイヤーを削除", en: "Remove empty layers{slash}sub-layers" },
            includeArtboardNumber: { ja: "アートボード番号", en: "Artboard Number" },
            includeArtboardName: { ja: "アートボード名", en: "Artboard Name" },
            useSeparator: { ja: "区切り文字", en: "Separator" }
        },
        panel: {
            exclude: { ja: "対象外にする", en: "Exclude" },
            target: { ja: "対象のアートボード", en: "Target Artboards" },
            layerName: { ja: "レイヤー名", en: "Layer Name" },
            locked: { ja: "ロック", en: "Locked" },
            hidden: { ja: "非表示", en: "Hidden" },
            layer: { ja: "レイヤー", en: "Layer" },
            object: { ja: "オブジェクト", en: "Object" },
            postProcess: { ja: "後処理", en: "Post-Processing" }
        },
        label: {
            specifiedLayers: { ja: "指定レイヤー", en: "Specified Layers" }
        },
        dropdown: {
            separatorUnderscore: { ja: "アンダースコア (_) ", en: "Underscore (_)" },
            separatorHyphen: { ja: "ハイフン (-)", en: "Hyphen (-)" },
            separatorSpace: { ja: "半角スペース", en: "Space" },
            separatorNone: { ja: "なし", en: "None" }
        },
        fallback: {
            artboard: { ja: "アートボード", en: "Artboard" }
        },
        radio: {
            currentArtboardOnly: { ja: "現在のアートボードのみ", en: "Current artboard" },
            allArtboards: { ja: "すべて", en: "All" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            failed: {
                ja: "一部のオブジェクトを移動できませんでした。\n移動失敗: ",
                en: "Some objects could not be moved.\nFailed moves: "
            }
        },
        hint: {
            excludedLayerNames: {
                ja: "カンマ または 「、」 区切りでレイヤー名を指定（例: bg, temp）",
                en: "Layer names separated by comma (e.g. bg, temp)"
            }
        }
    };

    /* LABELS からドット区切りのパスで多言語テキストを取り出す / Look up localized text from LABELS */
    function L(path) {
        var parts = path.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            node = node[parts[i]];
            if (!node) return path;
        }
        var text = node[lang] || node["en"];
        if (!text) return path;
        return applyUISymbols(text);
    }

    /* {slash} などのプレースホルダを言語別の記号に展開 / Expand placeholders like {slash} */
    function applyUISymbols(text) {
        return text
            .replace(/\{slash\}/g, uiSymbol("slash"))
            .replace(/\{colon\}/g, uiSymbol("colon"))
            .replace(/\{comma\}/g, uiSymbol("comma"))
            .replace(/\{openParen\}/g, uiSymbol("openParen"))
            .replace(/\{closeParen\}/g, uiSymbol("closeParen"));
    }

    /* 言語別の記号を返す（日本語＝全角、英語＝半角） / Return language-specific symbol */
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

    // =========================================
    // メイン / Main
    // =========================================

    if (app.documents.length === 0) {
        alert(L("alert.noDoc"));
        return;
    }

    var doc = app.activeDocument;
    var artboards = doc.artboards;
    var guideLayer = null; // _guide レイヤー参照（必要時に findLayerByName / getOrCreateGuideLayer で更新）

    /* パネル共通の見た目をまとめて設定 / Apply common panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ['fill', 'top'];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* ダイアログを構築・表示し、選択結果のオプションを返す / Build and show the options dialog */
    function showOptionsDialog() {
        var dlg = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];
        dlg.margins = [15, 20, 15, 15];

        var ctlTarget = buildTargetPanel(dlg);
        var ctlLayerName = buildLayerNamePanel(dlg);
        var ctlExclude = buildExcludePanel(dlg);
        var ctlPostProcess = buildPostProcessPanel(dlg);
        buildButtonRow(dlg);

        if (dlg.show() !== 1) {
            return null;
        }

        return {
            removeEmptyLayers: ctlPostProcess.chkRemoveEmpty.value,
            currentArtboardOnly: ctlTarget.rbCurrentArtboardOnly.value,
            includeArtboardNumber: ctlLayerName.chkIncludeArtboardNumber.value,
            includeArtboardName: ctlLayerName.chkIncludeArtboardName.value,
            layerNameSeparatorIndex: ctlLayerName.getSeparatorIndex(),
            ignoreLockedLayers: ctlExclude.chkIgnoreLockedLayers.value,
            ignoreLockedObjects: ctlExclude.chkIgnoreLockedObjects.value,
            ignoreHiddenLayers: ctlExclude.chkIgnoreHiddenLayers.value,
            ignoreHiddenObjects: ctlExclude.chkIgnoreHiddenObjects.value,
            excludedLayerNames: parseLayerNameList(ctlExclude.etSpecifiedLayers.text)
        };
    }

    /* 対象アートボードパネル（現在のみ／すべて） / Build the target-artboards panel */
    function buildTargetPanel(parent) {
        var panel = parent.add("panel", undefined, L("panel.target"));
        setupPanel(panel);

        var group = panel.add("group");
        group.orientation = "row";
        group.alignChildren = ["left", "center"];

        var rbCurrentArtboardOnly = group.add("radiobutton", undefined, L("radio.currentArtboardOnly"));
        var rbAllArtboards = group.add("radiobutton", undefined, L("radio.allArtboards"));

        if (artboards.length <= 1) {
            rbCurrentArtboardOnly.value = true;
            rbAllArtboards.enabled = false;
        } else {
            rbAllArtboards.value = true;
        }

        return {
            rbCurrentArtboardOnly: rbCurrentArtboardOnly,
            rbAllArtboards: rbAllArtboards
        };
    }

    /* レイヤー名パネル（番号・名前・区切り文字） / Build the layer-name panel */
    function buildLayerNamePanel(parent) {
        var panel = parent.add("panel", undefined, L("panel.layerName"));
        setupPanel(panel);

        var chkIncludeArtboardNumber = panel.add("checkbox", undefined, L("checkbox.includeArtboardNumber"));
        chkIncludeArtboardNumber.value = DEFAULTS.includeArtboardNumber;

        var grpSeparator = panel.add("group");
        grpSeparator.orientation = "row";
        grpSeparator.alignChildren = ["left", "center"];
        var chkUseSeparator = grpSeparator.add("checkbox", undefined, L("checkbox.useSeparator"));
        chkUseSeparator.value = DEFAULTS.useSeparator;
        var ddSeparator = grpSeparator.add("dropdownlist", undefined, [
            L("dropdown.separatorUnderscore"),
            L("dropdown.separatorHyphen"),
            L("dropdown.separatorSpace"),
            L("dropdown.separatorNone")
        ]);
        ddSeparator.selection = DEFAULTS.separatorIndex;

        function updateSeparatorEnabled() {
            var enabled = chkIncludeArtboardNumber.value;
            grpSeparator.enabled = enabled;
            ddSeparator.enabled = enabled && chkUseSeparator.value;
        }
        updateSeparatorEnabled();
        chkIncludeArtboardNumber.onClick = updateSeparatorEnabled;
        chkUseSeparator.onClick = updateSeparatorEnabled;

        var chkIncludeArtboardName = panel.add("checkbox", undefined, L("checkbox.includeArtboardName"));
        chkIncludeArtboardName.value = DEFAULTS.includeArtboardName;

        return {
            chkIncludeArtboardNumber: chkIncludeArtboardNumber,
            chkIncludeArtboardName: chkIncludeArtboardName,
            getSeparatorIndex: function () {
                return (chkUseSeparator.value && ddSeparator.selection) ? ddSeparator.selection.index : 3;
            }
        };
    }

    /* 対象外パネル（ロック／非表示／指定レイヤー） / Build the exclude panel */
    function buildExcludePanel(parent) {
        var panel = parent.add("panel", undefined, L("panel.exclude"));
        setupPanel(panel);

        var grpLockedHidden = panel.add("group");
        grpLockedHidden.orientation = "row";
        grpLockedHidden.alignChildren = ["left", "top"];
        grpLockedHidden.spacing = 15;

        var lockedSub = buildLockedHiddenSubPanel(grpLockedHidden, "panel.locked", DEFAULTS.ignoreLockedLayers, DEFAULTS.ignoreLockedObjects);
        var hiddenSub = buildLockedHiddenSubPanel(grpLockedHidden, "panel.hidden", DEFAULTS.ignoreHiddenLayers, DEFAULTS.ignoreHiddenObjects);

        /* 指定レイヤー入力欄（対象外パネル最下部） / Specified layer names input */
        var grpSpecified = panel.add("group");
        grpSpecified.orientation = "row";
        grpSpecified.alignChildren = ["left", "center"];
        grpSpecified.add("statictext", undefined, L("label.specifiedLayers") + uiSymbol("colon"));
        var etSpecifiedLayers = grpSpecified.add("edittext", undefined, DEFAULTS.excludedLayerNames);
        etSpecifiedLayers.preferredSize.width = 200;
        etSpecifiedLayers.helpTip = L("hint.excludedLayerNames");

        return {
            chkIgnoreLockedLayers: lockedSub.chkLayer,
            chkIgnoreLockedObjects: lockedSub.chkObject,
            chkIgnoreHiddenLayers: hiddenSub.chkLayer,
            chkIgnoreHiddenObjects: hiddenSub.chkObject,
            etSpecifiedLayers: etSpecifiedLayers
        };
    }

    /* ロック／非表示サブパネル（共通レイアウト） / Build a locked-or-hidden sub-panel */
    function buildLockedHiddenSubPanel(parent, titleKey, defaultLayerValue, defaultObjectValue) {
        var panel = parent.add("panel", undefined, L(titleKey));
        setupPanel(panel);

        var chkLayer = panel.add("checkbox", undefined, L("panel.layer"));
        chkLayer.value = defaultLayerValue;
        var chkObject = panel.add("checkbox", undefined, L("panel.object"));
        chkObject.value = defaultObjectValue;

        return { chkLayer: chkLayer, chkObject: chkObject };
    }

    /* 後処理パネル（空レイヤー削除） / Build the post-processing panel */
    function buildPostProcessPanel(parent) {
        var panel = parent.add("panel", undefined, L("panel.postProcess"));
        setupPanel(panel);
        var chkRemoveEmpty = panel.add("checkbox", undefined, L("checkbox.removeEmpty"));
        chkRemoveEmpty.value = DEFAULTS.removeEmptyLayers;
        return { chkRemoveEmpty: chkRemoveEmpty };
    }

    /* OK / Cancel ボタン列 / Build the OK / Cancel button row */
    function buildButtonRow(dlg) {
        var btnGroup = dlg.add("group");
        btnGroup.orientation = "row";
        btnGroup.alignment = ["center", "top"];
        var btnCancel = btnGroup.add("button", undefined, L("button.cancel"), { name: "cancel" });
        var btnOk = btnGroup.add("button", undefined, L("button.ok"), { name: "ok" });
        dlg.defaultElement = btnOk;
        dlg.cancelElement = btnCancel;
    }

    /* "bg, temp" のような文字列を配列に分解（,または、区切り） / Parse comma-separated list into array */
    function parseLayerNameList(text) {
        if (!text) return [];
        var parts = text.split(/[,、]/);
        var result = [];
        for (var i = 0; i < parts.length; i++) {
            var trimmed = parts[i].replace(/^\s+|\s+$/g, "");
            if (trimmed.length > 0) result.push(trimmed);
        }
        return result;
    }

    /* レイヤー名が指定除外リストに含まれるか / Check if layer name is in the excluded list */
    function isLayerNameExcluded(layerName, list) {
        if (!list || list.length === 0) return false;
        for (var i = 0; i < list.length; i++) {
            if (list[i] === layerName) return true;
        }
        return false;
    }

    /* レイヤーが除外対象か判定（指定名・ロック・非表示） / Check if a layer should be ignored */
    function shouldIgnoreLayer(layer, options) {
        if (!layer || layer.typename !== "Layer") return false;
        if (isLayerNameExcluded(layer.name, options.excludedLayerNames)) return true;
        if (options.ignoreLockedLayers && layer.locked) return true;
        if (options.ignoreHiddenLayers && !layer.visible) return true;
        return false;
    }

    /* 親方向に辿って除外対象のレイヤーが存在するか / Check if any ancestor layer is ignored */
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

    /* 親方向に辿って「ロック／非表示」だけを理由に除外される祖先があるか（指定レイヤー名は無視）
       Check ancestors for lock/hidden exclusion only (skip name-based exclusion) */
    function hasIgnoredAncestorByLockOrHidden(obj, options) {
        var parent = obj.parent;
        while (parent) {
            if (parent.typename === "Layer") {
                if (options.ignoreLockedLayers && parent.locked) return true;
                if (options.ignoreHiddenLayers && !parent.visible) return true;
            }
            if (parent.typename === "Document") {
                break;
            }
            parent = parent.parent;
        }
        return false;
    }

    /* オブジェクトが除外対象か判定（ロック／非表示） / Check if an object should be ignored
       PageItem の非表示判定は visible ではなく hidden を使う
       一部の PageItem は locked / hidden 参照で例外になることがあるため、各プロパティを個別に判定 */
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

    /* 名前でトップレベルレイヤーを検索 / Find a top-level layer by name */
    function findLayerByName(name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) {
                return doc.layers[i];
            }
        }
        return null;
    }

    /* 指定名のレイヤーを取得、無ければ作成 / Get a top-level layer by name, creating it if missing */
    function getOrCreateLayer(name) {
        var layer = findLayerByName(name);
        if (layer) {
            return layer;
        }
        layer = doc.layers.add();
        layer.name = name;
        return layer;
    }

    /* _guide レイヤーを必要時に取得・作成 / Get or create the _guide layer on demand */
    function getOrCreateGuideLayer() {
        if (!guideLayer) {
            guideLayer = getOrCreateLayer(GUIDE_LAYER_NAME);
        }
        return guideLayer;
    }

    /* システム管理レイヤー（_guide / _pasteboard）かどうか / Is this a script-managed system layer */
    function isProtectedSystemLayer(layer) {
        if (!layer || layer.typename !== "Layer") return false;
        return layer.name === GUIDE_LAYER_NAME || layer.name === PASTEBOARD_LAYER_NAME;
    }

    /* オプション設定に応じたレイヤー名区切り文字を返す / Return the separator string for layer names */
    function getLayerNameSeparator(options) {
        var index = options && typeof options.layerNameSeparatorIndex === "number" ? options.layerNameSeparatorIndex : 0;
        switch (index) {
            case 1: return "-";
            case 2: return " ";
            case 3: return "";
            default: return "_";
        }
    }

    /* アートボード番号と名前からレイヤー名を組み立て / Build a layer name for the given artboard index */
    function getArtboardLayerName(index, options) {
        var parts = [];
        var artboardName = artboards[index].name;
        if (!artboardName || artboardName === "") {
            artboardName = L("fallback.artboard");
        }

        if (!options || options.includeArtboardNumber !== false) {
            parts.push(String(index + 1));
        }
        if (!options || options.includeArtboardName !== false) {
            parts.push(artboardName);
        }

        if (parts.length === 0) {
            parts.push(String(index + 1));
            parts.push(artboardName);
        }
        return parts.join(getLayerNameSeparator(options));
    }

    /* 旧仕様レイヤー名（アートボード名のみ）からアートボード index を引く / Find artboard index from a legacy layer name */
    function findLegacyArtboardIndexByLayerName(name) {
        for (var i = 0; i < artboards.length; i++) {
            var artboardName = artboards[i].name;
            if (!artboardName || artboardName === "") {
                artboardName = L("fallback.artboard");
            }
            if (artboardName === name) {
                return i;
            }
        }
        return -1;
    }

    /* レイヤーを一時的に書き込み可（ロック解除・表示）にして fn を実行し、終了時に元の状態へ戻す
       fn 内でレイヤーが削除されていた場合は復元せず黙って抜ける
       Run fn with the layer temporarily writable; restore state afterwards (tolerate removal inside fn) */
    function withWritableLayer(layer, fn) {
        if (!layer || layer.typename !== "Layer") {
            return fn();
        }
        var wasLocked = layer.locked;
        var wasVisible = layer.visible;
        if (wasLocked) layer.locked = false;
        if (!wasVisible) layer.visible = true;
        try {
            return fn();
        } finally {
            try {
                if (wasLocked) layer.locked = true;
                if (!wasVisible) layer.visible = false;
            } catch (e) {
                // fn 内でレイヤー削除済み等で参照不能ならスキップ
            }
        }
    }

    /* 収集したエントリ群を移動先レイヤーへ移動 / Move collected entries into the target layer */
    function moveCollectedEntries(entries, targetLayer, processed) {
        var result = {
            moved: 0,
            failed: 0
        };
        if (entries.length === 0) return result;
        return withWritableLayer(targetLayer, function () {
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
        });
    }

    /* レイヤー以下の pageItem を再帰的に集める / Recursively collect pageItems from a layer */
    function collectLayerPageItemsRecursive(layer, outEntries) {
        for (var i = 0; i < layer.layers.length; i++) {
            collectLayerPageItemsRecursive(layer.layers[i], outEntries);
        }
        for (var j = 0; j < layer.pageItems.length; j++) {
            var pageItem = layer.pageItems[j];
            if (pageItem.parent !== layer) continue;
            outEntries.push({ item: pageItem });
        }
    }

    /* レイヤー内の全 pageItem を別レイヤーへ移動 / Move all pageItems from a source layer into target */
    function moveLayerItemsToLayer(sourceLayer, targetLayer) {
        var entries = [];
        collectLayerPageItemsRecursive(sourceLayer, entries);
        return moveCollectedEntries(entries, targetLayer, null);
    }

    /* 空になったサブレイヤーを再帰的に削除 / Remove empty sub-layers recursively */
    function removeEmptySubLayers(parentLayer, options) {
        for (var i = parentLayer.layers.length - 1; i >= 0; i--) {
            var subLayer = parentLayer.layers[i];
            removeEmptySubLayers(subLayer, options);
            if (shouldIgnoreLayer(subLayer, options)) continue;
            if (subLayer.pageItems.length === 0 && subLayer.layers.length === 0) {
                subLayer.remove();
            }
        }
    }

    /* オブジェクトの重心がアートボード矩形に含まれるか / Check if item's centroid is within artboard rect */
    function isInArtboard(item, rect) {
        var bounds = item.geometricBounds; // [left, top, right, bottom]
        var centroidX = (bounds[0] + bounds[2]) / 2;
        var centroidY = (bounds[1] + bounds[3]) / 2;

        return (
            centroidX >= rect[0] &&
            centroidX <= rect[2] &&
            centroidY <= rect[1] &&
            centroidY >= rect[3]
        );
    }

    /* トップレベルの対象 pageItems を集める。ガイドは指定レイヤー除外を無視して常に回収する
       Collect target top-level pageItems; guides bypass name-based exclusion */
    function collectTargetItems(options) {
        var items = [];
        for (var i = 0; i < doc.pageItems.length; i++) {
            var pageItem = doc.pageItems[i];
            var parentNode = pageItem.parent;
            if (parentNode.typename !== "Layer" && parentNode.typename !== "Document") continue;

            var isGuide = false;
            try { isGuide = pageItem.guides; } catch (e) {}

            if (isGuide) {
                // ガイドは _guide へ集約するため、指定レイヤー名による除外は無視。
                // ただしロック／非表示の祖先（移動できない）は尊重する。
                if (hasIgnoredAncestorByLockOrHidden(pageItem, options)) continue;
                items.push(pageItem);
                continue;
            }

            if (hasIgnoredAncestorLayer(pageItem, options)) continue;
            if (shouldIgnoreObject(pageItem, options)) continue;
            items.push(pageItem);
        }
        return items;
    }

    /* 処理対象のアートボード範囲（start <= idx < end）を返す / Return artboard index range to process */
    function getArtboardRange(options) {
        if (options.currentArtboardOnly) {
            var start = doc.artboards.getActiveArtboardIndex();
            return { start: start, end: start + 1 };
        }
        return { start: 0, end: artboards.length };
    }

    /* 各アートボードに対応するレイヤーへアイテムを振り分け / Assign items to each artboard's layer */
    function assignItemsToArtboardLayers(items, options, processed, range) {
        var failedMoves = 0;
        for (var artboardIdx = range.start; artboardIdx < range.end; artboardIdx++) {
            var artboardLayer = getOrCreateLayer(getArtboardLayerName(artboardIdx, options));
            var artboardRect = artboards[artboardIdx].artboardRect;
            var normalEntries = [];
            var guideEntries = [];

            for (var i = 0; i < items.length; i++) {
                if (processed[i]) continue;
                var candidateItem = items[i];
                try {
                    if (!isInArtboard(candidateItem, artboardRect)) continue;
                    if (candidateItem.guides) {
                        guideEntries.push({ item: candidateItem, index: i });
                    } else {
                        normalEntries.push({ item: candidateItem, index: i });
                    }
                } catch (e) {
                    // 判定できないものは無視
                }
            }

            failedMoves += moveCollectedEntries(normalEntries, artboardLayer, processed).failed;
            if (guideEntries.length > 0) {
                failedMoves += moveCollectedEntries(guideEntries, getOrCreateGuideLayer(), processed).failed;
            }
        }
        return failedMoves;
    }

    /* 対象アートボードのレイヤーを上から 1→2→3… の順に並べる / Reorder artboard layers to the top */
    function reorderArtboardLayersToFront(options, range) {
        for (var i = range.end - 1; i >= range.start; i--) {
            var layer = getOrCreateLayer(getArtboardLayerName(i, options));
            withWritableLayer(layer, function () {
                layer.zOrder(ZOrderMethod.BRINGTOFRONT);
            });
        }
    }

    /* どのアートボードにも属さないアイテムを _pasteboard / _guide に振り分け / Assign remaining items to pasteboard or guide */
    function assignPasteboardItems(items, options, processed) {
        var failedMoves = 0;
        var pasteboardNormalEntries = [];
        var pasteboardGuideEntries = [];
        for (var i = 0; i < items.length; i++) {
            if (processed[i]) continue;
            if (items[i].guides) {
                pasteboardGuideEntries.push({ item: items[i], index: i });
            } else {
                pasteboardNormalEntries.push({ item: items[i], index: i });
            }
        }
        if (pasteboardNormalEntries.length > 0) {
            failedMoves += moveCollectedEntries(pasteboardNormalEntries, getOrCreateLayer(PASTEBOARD_LAYER_NAME), processed).failed;
        }
        if (pasteboardGuideEntries.length > 0) {
            failedMoves += moveCollectedEntries(pasteboardGuideEntries, getOrCreateGuideLayer(), processed).failed;
        }
        return failedMoves;
    }

    /* _guide レイヤーが存在すれば最前面へ / Bring _guide to front if it exists */
    function bringGuideLayerToFront() {
        if (!guideLayer) return;
        withWritableLayer(guideLayer, function () {
            guideLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        });
    }

    /* 旧仕様（アートボード名のみ）レイヤーを新仕様レイヤーへ統合 / Merge legacy artboard-named layers into new layers */
    function mergeLegacyLayers(options) {
        var failedMoves = 0;
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            var legacyLayer = doc.layers[i];
            if (isProtectedSystemLayer(legacyLayer)) continue;
            if (shouldIgnoreLayer(legacyLayer, options)) continue;
            var legacyArtboardIndex = findLegacyArtboardIndexByLayerName(legacyLayer.name);
            if (legacyArtboardIndex < 0) continue;

            var targetLayer = getOrCreateLayer(getArtboardLayerName(legacyArtboardIndex, options));
            if (legacyLayer === targetLayer) continue;

            withWritableLayer(legacyLayer, function () {
                failedMoves += moveLayerItemsToLayer(legacyLayer, targetLayer).failed;
                if (legacyLayer.pageItems.length === 0 && legacyLayer.layers.length === 0 && doc.layers.length > 1) {
                    legacyLayer.remove();
                }
            });
        }
        return failedMoves;
    }

    /* 処理後の空レイヤー削除（_guide / _pasteboard は保護） / Cleanup empty layers (protect system layers) */
    function cleanupEmptyLayers(options) {
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            var topLayer = doc.layers[i];
            if (shouldIgnoreLayer(topLayer, options)) continue;
            removeEmptySubLayers(topLayer, options);
            if (isProtectedSystemLayer(topLayer)) continue;
            if (topLayer.pageItems.length === 0 && topLayer.layers.length === 0 && doc.layers.length > 1) {
                topLayer.remove();
            }
        }
    }

    // -- ここから本処理 / Begin main flow --

    var options = showOptionsDialog();
    if (!options) {
        return;
    }

    var items = collectTargetItems(options);
    guideLayer = findLayerByName(GUIDE_LAYER_NAME); // 既存があれば先に拾う / capture existing if any
    var processed = [];
    var failedMoves = 0;
    var range = getArtboardRange(options);

    failedMoves += assignItemsToArtboardLayers(items, options, processed, range);
    reorderArtboardLayersToFront(options, range);

    if (!options.currentArtboardOnly) {
        failedMoves += assignPasteboardItems(items, options, processed);
    }

    bringGuideLayerToFront();

    if (!options.currentArtboardOnly) {
        failedMoves += mergeLegacyLayers(options);
    }

    reorderArtboardLayersToFront(options, range);
    bringGuideLayerToFront();

    if (options.removeEmptyLayers) {
        cleanupEmptyLayers(options);
    }

    if (failedMoves > 0) {
        alert(L("alert.failed") + failedMoves);
    }

})();
