#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

    (function () {

    // =========================================
    // バージョンとローカライズ / Version and Localization
    // =========================================

    var SCRIPT_VERSION = "v1.0.0";

    /* 現在のロケールを判定 / Detect current locale */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    var LABELS = {
        dialogTitle:         { ja: "名前の検索置換", en: "Find and Replace Names" },
        noDoc:               { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        target:              { ja: "対象", en: "Target" },
        artboard:            { ja: "アートボード", en: "Artboard" },
        layer:               { ja: "レイヤー", en: "Layer" },
        symbol:              { ja: "シンボル", en: "Symbol" },
        graphicStyle:        { ja: "グラフィックスタイル", en: "Graphic Style" },
        findReplace:         { ja: "検索・置換", en: "Find & Replace" },
        findReplaceEnable:   { ja: "検索置換", en: "Find & Replace" },
        find:                { ja: "検索", en: "Find" },
        replace:             { ja: "置換", en: "Replace" },
        regex:               { ja: "正規表現", en: "Regex" },
        prefix:              { ja: "接頭辞", en: "Prefix" },
        suffix:              { ja: "接尾辞", en: "Suffix" },
        numberingEnable:     { ja: "ナンバリング", en: "Numbering" },
        separator:           { ja: "区切り:", en: "Separator:" },
        startNumber:         { ja: "開始番号:", en: "Start:" },
        sort:                { ja: "並び替え:", en: "Sort:" },
        sortOriginal:        { ja: "元の順", en: "Original" },
        sortNameAsc:         { ja: "名前 ↑", en: "Name ↑" },
        sortNameDesc:        { ja: "名前 ↓", en: "Name ↓" },
        sortChanged:         { ja: "変更あり優先", en: "Changed first" },
        moveTop:             { ja: "↑↑", en: "↑↑" },
        moveUp:              { ja: "↑", en: "↑" },
        moveDown:            { ja: "↓", en: "↓" },
        moveBottom:          { ja: "↓↓", en: "↓↓" },
        cancel:              { ja: "キャンセル", en: "Cancel" },
        needInput:           { ja: "検索文字を入力するか、接頭辞・接尾辞のナンバリングを有効にしてください。", en: "Enter a search string or enable prefix/suffix numbering." },
        noMatchArtboard:     { ja: "該当するアートボード名はありませんでした。", en: "No artboard names matched." },
        noMatchLayer:        { ja: "該当するレイヤー名はありませんでした。", en: "No layer names matched." },
        noMatchSymbol:       { ja: "該当するシンボル名はありませんでした。", en: "No symbol names matched." },
        noMatchGraphicStyle: { ja: "該当するグラフィックスタイル名はありませんでした。", en: "No graphic style names matched." },
        done:                { ja: "完了しました。", en: "Done." },
        targetArtboards:     { ja: "対象: アートボード名", en: "Target: Artboards" },
        targetLayers:        { ja: "対象: レイヤー名", en: "Target: Layers" },
        targetSymbols:       { ja: "対象: シンボル名", en: "Target: Symbols" },
        targetGraphicStyles: { ja: "対象: グラフィックスタイル名", en: "Target: Graphic Styles" },
        renamed:             { ja: "変更数: ", en: "Renamed: " },
        suffixed:            { ja: "同名回避で連番追加: ", en: "Suffixed to avoid duplicates: " },
        errorsLabel:         { ja: "エラー: ", en: "Errors: " },
        countSuffix:         { ja: " 件", en: "" }
    };

    /* ローカライズ文字列を取得 / Get localized string */
    function L(key) {
        var entry = LABELS[key];
        if (!entry) return key;
        return entry[lang] || entry.en || entry.ja || key;
    }

        if (app.documents.length === 0) {
            alert(L("noDoc"));
            return;
        }

        var doc = app.activeDocument;

        // =========================================
        // ダイアログ定数 / ヘルパー / Dialog constants & helpers
        // =========================================

        var PANEL_MARGINS = [15, 20, 15, 10];
        var PANEL_SPACING = 8;
        var PREVIEW_LINE_HEIGHT = 16; // Mac の listbox 行高さ目安 / Approx. line height on Mac
        var PREVIEW_VISIBLE_LINES = 20;

        /* パネルの共通設定を適用 / Apply common panel settings */
        function setupPanel(panel, spacing) {
            panel.orientation = "column";
            panel.alignChildren = "left";
            panel.alignment = "fill";
            panel.margins = PANEL_MARGINS;
            panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
        }

        // =========================================
        // 共通関数 / Common functions
        // =========================================

        /* 全置換（非正規表現） / Replace all (non-regex) */
        function replaceAll(text, search, replacement) {
            return text.split(search).join(replacement);
        }

        /* Object.prototype 衝突回避のキー / Key to avoid Object.prototype collisions */
        function nameKey(name) {
            return "@" + name;
        }

        /* 衝突しないユニーク名を生成 / Generate unique name avoiding collisions */
        function makeUniqueName(baseName, usedNames) {
            var name = baseName;
            var suffix = 2;

            while (usedNames[nameKey(name)]) {
                name = baseName + "_" + suffix;
                suffix++;
            }

            return name;
        }

        /* 名前取得の例外を握りつぶす / Safely get item name */
        function safeGetName(item) {
            try { return item.name; } catch (e) { return ""; }
        }

        /* 検索文字 / 正規表現でマッチ判定 / Test match by string or regex */
        function matchesFind(text, search, useRegex) {
            if (search === "") return false;
            if (useRegex) {
                try { return new RegExp(search).test(text); }
                catch (e) { return false; }
            }
            return text.indexOf(search) >= 0;
        }

        /* 置換を実行 / Apply replacement */
        function applyReplace(text, search, replacement, useRegex) {
            if (useRegex) {
                try { return text.replace(new RegExp(search, "g"), replacement); }
                catch (e) { return text; }
            }
            return replaceAll(text, search, replacement);
        }

        /* 0埋め / Zero-pad number */
        function padLeftZero(num, width) {
            var s = "" + num;
            while (s.length < width) s = "0" + s;
            return s;
        }

        /* 入力文字列から開始番号と桁数を取得 / Parse start number and width */
        function parseStartSpec(s) {
            var n = parseInt(s, 10);
            if (isNaN(n)) n = 1;
            return { start: n, width: s.length > 0 ? s.length : 1 };
        }

        /* ネストレイヤーを再帰的に収集 / Recursively collect nested layers */
        function collectLayers(layers, collected) {
            for (var i = 0; i < layers.length; i++) {
                var layer = layers[i];

                collected.push(layer);

                if (layer.layers && layer.layers.length > 0) {
                    collectLayers(layer.layers, collected);
                }
            }
        }

        /* 対象モードに応じた項目配列を返す / Return items for the given mode */
        function getTargetItems(mode) {
            var items = [];
            var i;
            if (mode === "artboard") {
                for (i = 0; i < doc.artboards.length; i++) items.push(doc.artboards[i]);
            } else if (mode === "layer") {
                collectLayers(doc.layers, items);
            } else if (mode === "symbol") {
                for (i = 0; i < doc.symbols.length; i++) items.push(doc.symbols[i]);
            } else if (mode === "graphicStyle") {
                for (i = 0; i < doc.graphicStyles.length; i++) items.push(doc.graphicStyles[i]);
            }
            return items;
        }

        /* ラジオボタンから対象モードを取得 / Get current target mode from radios */
        function getModeFromRadios() {
            if (rbArtboard.value) return "artboard";
            if (rbLayer.value) return "layer";
            if (rbSymbol.value) return "symbol";
            if (rbGraphicStyle.value) return "graphicStyle";
            return "symbol";
        }

        // =========================================
        // ダイアログ / Dialog
        // =========================================

        var win = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        win.orientation = "column";
        win.alignChildren = ["fill", "top"];

        // 対象選択 (2カラムを貫通) / Target selection (spans both columns)
        var targetPanel = win.add("panel", undefined, L("target"));
        targetPanel.orientation = "row";
        targetPanel.alignChildren = ["left", "center"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = 10;

        var rbArtboard = targetPanel.add("radiobutton", undefined, L("artboard"));
        var rbLayer = targetPanel.add("radiobutton", undefined, L("layer"));
        var rbSymbol = targetPanel.add("radiobutton", undefined, L("symbol"));
        var rbGraphicStyle = targetPanel.add("radiobutton", undefined, L("graphicStyle"));

        rbSymbol.value = true;

        // 2カラム / Two-column layout
        var mainGroup = win.add("group");
        mainGroup.orientation = "row";
        mainGroup.alignChildren = ["fill", "fill"];

        // 左カラム / Left column
        var leftColumn = mainGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];

        // 検索・置換パネル / Find & Replace panel
        var findReplacePanel = leftColumn.add("panel", undefined, L("findReplace"));
        setupPanel(findReplacePanel, 6);

        var findReplaceCheckbox = findReplacePanel.add("checkbox", undefined, L("findReplaceEnable"));
        findReplaceCheckbox.value = true;

        var findGroup = findReplacePanel.add("group");
        findGroup.orientation = "row";
        findGroup.alignChildren = ["left", "center"];
        findGroup.add("statictext", undefined, L("find"));
        var findInput = findGroup.add("edittext", undefined, "");
        findInput.characters = 15;

        var replaceGroup = findReplacePanel.add("group");
        replaceGroup.orientation = "row";
        replaceGroup.alignChildren = ["left", "center"];
        replaceGroup.add("statictext", undefined, L("replace"));
        var replaceInput = replaceGroup.add("edittext", undefined, "");
        replaceInput.characters = 15;

        var regexGroup = findReplacePanel.add("group");
        regexGroup.orientation = "row";
        regexGroup.alignChildren = ["right", "center"];
        regexGroup.alignment = "fill";
        var regexCheckbox = regexGroup.add("checkbox", undefined, L("regex"));

        // 接頭辞パネル / Prefix panel
        var prefixPanel = leftColumn.add("panel", undefined, L("prefix"));
        setupPanel(prefixPanel, 6);

        var prefixCheckbox = prefixPanel.add("checkbox", undefined, L("numberingEnable"));

        var prefixSepGroup = prefixPanel.add("group");
        prefixSepGroup.orientation = "row";
        prefixSepGroup.alignChildren = ["left", "center"];
        prefixSepGroup.add("statictext", undefined, L("separator"));
        var rbPrefixSepDash = prefixSepGroup.add("radiobutton", undefined, "-");
        var rbPrefixSepUnderscore = prefixSepGroup.add("radiobutton", undefined, "_");
        rbPrefixSepDash.value = true;

        var prefixStartGroup = prefixPanel.add("group");
        prefixStartGroup.orientation = "row";
        prefixStartGroup.alignChildren = ["left", "center"];
        prefixStartGroup.add("statictext", undefined, L("startNumber"));
        var prefixStartInput = prefixStartGroup.add("edittext", undefined, "1");
        prefixStartInput.characters = 4;

        // 接尾辞パネル / Suffix panel
        var suffixPanel = leftColumn.add("panel", undefined, L("suffix"));
        setupPanel(suffixPanel, 6);

        var suffixCheckbox = suffixPanel.add("checkbox", undefined, L("numberingEnable"));

        var suffixSepGroup = suffixPanel.add("group");
        suffixSepGroup.orientation = "row";
        suffixSepGroup.alignChildren = ["left", "center"];
        suffixSepGroup.add("statictext", undefined, L("separator"));
        var rbSuffixSepDash = suffixSepGroup.add("radiobutton", undefined, "-");
        var rbSuffixSepUnderscore = suffixSepGroup.add("radiobutton", undefined, "_");
        rbSuffixSepDash.value = true;

        var suffixStartGroup = suffixPanel.add("group");
        suffixStartGroup.orientation = "row";
        suffixStartGroup.alignChildren = ["left", "center"];
        suffixStartGroup.add("statictext", undefined, L("startNumber"));
        var suffixStartInput = suffixStartGroup.add("edittext", undefined, "1");
        suffixStartInput.characters = 4;

        // 右カラム: 並び替え + プレビュー / Right column: sort + preview
        var rightColumn = mainGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "fill"];

        var sortGroup = rightColumn.add("group");
        sortGroup.orientation = "row";
        sortGroup.alignChildren = ["left", "center"];
        sortGroup.add("statictext", undefined, L("sort"));
        var sortDropdown = sortGroup.add("dropdownlist", undefined, [
            L("sortOriginal"),
            L("sortNameAsc"),
            L("sortNameDesc"),
            L("sortChanged")
        ]);
        sortDropdown.selection = 0;

        var moveTopBtn = sortGroup.add("button", undefined, L("moveTop"));
        var moveUpBtn = sortGroup.add("button", undefined, L("moveUp"));
        var moveDownBtn = sortGroup.add("button", undefined, L("moveDown"));
        var moveBottomBtn = sortGroup.add("button", undefined, L("moveBottom"));
        moveTopBtn.preferredSize.width = 36;
        moveUpBtn.preferredSize.width = 32;
        moveDownBtn.preferredSize.width = 32;
        moveBottomBtn.preferredSize.width = 36;

        var previewList = rightColumn.add("listbox", undefined, []);
        previewList.preferredSize.width = 340;
        previewList.preferredSize.height = PREVIEW_VISIBLE_LINES * PREVIEW_LINE_HEIGHT;

        // ボタン (Mac規約: Cancel → OK) / Buttons (Mac convention)
        var buttonGroup = win.add("group");
        buttonGroup.alignment = "right";

        buttonGroup.add("button", undefined, L("cancel"), { name: "cancel" });
        var okBtn = buttonGroup.add("button", undefined, "OK", { name: "ok" });

        okBtn.enabled = false;

        // =========================================
        // プレビュー / ナンバリング / Preview & numbering
        // =========================================

        var lastItems = [];

        /* 区切り文字を取得 (接頭辞/接尾辞) / Get separator for prefix/suffix */
        function getPrefixSeparator() {
            return rbPrefixSepDash.value ? "-" : "_";
        }
        function getSuffixSeparator() {
            return rbSuffixSepDash.value ? "-" : "_";
        }

        /* 新名前を全件再計算 / Recompute all new names */
        function recomputeNewNames() {
            var prefixOn = prefixCheckbox.value;
            var suffixOn = suffixCheckbox.value;

            var prefixSep = getPrefixSeparator();
            var prefixSpec = parseStartSpec(prefixStartInput.text);
            var prefixCounter = prefixSpec.start;

            var suffixSep = getSuffixSeparator();
            var suffixSpec = parseStartSpec(suffixStartInput.text);
            var suffixCounter = suffixSpec.start;

            for (var i = 0; i < lastItems.length; i++) {
                var e = lastItems[i];
                var base = e.matched ? e.replacedName : e.oldName;
                var finalName = base;

                if (prefixOn) {
                    finalName = padLeftZero(prefixCounter, prefixSpec.width) + prefixSep + finalName;
                    prefixCounter++;
                }
                if (suffixOn) {
                    finalName = finalName + suffixSep + padLeftZero(suffixCounter, suffixSpec.width);
                    suffixCounter++;
                }

                e.newName = finalName;
            }
        }

        /* 並び替えを適用 / Apply sort */
        function applySort() {
            var mode = sortDropdown.selection ? sortDropdown.selection.index : 0;
            if (mode === 1) {
                lastItems.sort(function (a, b) {
                    return a.oldName < b.oldName ? -1 : (a.oldName > b.oldName ? 1 : 0);
                });
            } else if (mode === 2) {
                lastItems.sort(function (a, b) {
                    return a.oldName < b.oldName ? 1 : (a.oldName > b.oldName ? -1 : 0);
                });
            } else if (mode === 3) {
                lastItems.sort(function (a, b) {
                    return (b.matched ? 1 : 0) - (a.matched ? 1 : 0);
                });
            }
        }

        /* プレビュー (listbox) を描画 / Render preview list */
        function renderPreview(keepSelectionIndex) {
            previewList.removeAll();
            for (var i = 0; i < lastItems.length; i++) {
                var e = lastItems[i];
                var text;
                if (e.newName !== e.oldName) {
                    text = e.oldName + "  →  " + e.newName;
                } else {
                    text = e.oldName;
                }
                previewList.add("item", text);
            }
            if (typeof keepSelectionIndex === "number" && keepSelectionIndex >= 0 && keepSelectionIndex < previewList.items.length) {
                previewList.selection = keepSelectionIndex;
            }
        }

        /* 対象を再走査してプレビューを更新 / Rescan items and refresh preview */
        function refreshPreview() {
            var mode = getModeFromRadios();
            var items = getTargetItems(mode);
            var findReplaceOn = findReplaceCheckbox.value;
            var search = findInput.text;
            var replacement = replaceInput.text;
            var useRegexNow = regexCheckbox.value;

            lastItems = [];
            for (var i = 0; i < items.length; i++) {
                var oldName = safeGetName(items[i]);
                if (oldName === "") continue;

                var replaced = oldName;
                var matched = false;
                if (findReplaceOn && search !== "" && matchesFind(oldName, search, useRegexNow)) {
                    var candidate = applyReplace(oldName, search, replacement, useRegexNow);
                    if (candidate !== oldName) {
                        replaced = candidate;
                        matched = true;
                    }
                }

                lastItems.push({
                    item: items[i],
                    oldName: oldName,
                    replacedName: replaced,
                    matched: matched,
                    newName: replaced
                });
            }

            applySort();
            recomputeNewNames();
            renderPreview();
        }

        /* 再計算と再描画のみ実行 / Recompute & re-render without rescan */
        function rerender() {
            recomputeNewNames();
            var selIdx = previewList.selection ? previewList.selection.index : -1;
            renderPreview(selIdx);
        }

        /* 検索置換 UI の有効/無効を切替 / Toggle find & replace UI */
        function updateFindReplaceEnabled() {
            var on = findReplaceCheckbox.value;
            findInput.enabled = on;
            replaceInput.enabled = on;
            regexCheckbox.enabled = on;
        }

        /* 接頭辞/接尾辞 UI の有効/無効を切替 / Toggle prefix/suffix UI */
        function updatePrefixEnabled() {
            var on = prefixCheckbox.value;
            rbPrefixSepDash.enabled = on;
            rbPrefixSepUnderscore.enabled = on;
            prefixStartInput.enabled = on;
        }
        function updateSuffixEnabled() {
            var on = suffixCheckbox.value;
            rbSuffixSepDash.enabled = on;
            rbSuffixSepUnderscore.enabled = on;
            suffixStartInput.enabled = on;
        }

        /* OK ボタンの有効/無効を切替 / Toggle OK button */
        function updateOkEnabled() {
            var findReplaceActive = findReplaceCheckbox.value && findInput.text.length > 0;
            okBtn.enabled = findReplaceActive || prefixCheckbox.value || suffixCheckbox.value;
        }

        /* 選択行を上下に動かす / Move selected row up or down */
        function moveSelected(direction) {
            var sel = previewList.selection;
            if (!sel) return;
            var idx = sel.index;
            var newIdx = idx + direction;
            if (newIdx < 0 || newIdx >= lastItems.length) return;

            var tmp = lastItems[idx];
            lastItems[idx] = lastItems[newIdx];
            lastItems[newIdx] = tmp;

            recomputeNewNames();
            renderPreview(newIdx);
        }

        sortDropdown.onChange = function () {
            applySort();
            recomputeNewNames();
            renderPreview();
        };

        moveUpBtn.onClick = function () { moveSelected(-1); };
        moveDownBtn.onClick = function () { moveSelected(1); };

        rbArtboard.onClick = refreshPreview;
        rbLayer.onClick = refreshPreview;
        rbSymbol.onClick = refreshPreview;
        rbGraphicStyle.onClick = refreshPreview;

        findInput.onChanging = function () {
            updateOkEnabled();
            refreshPreview();
        };

        replaceInput.onChanging = refreshPreview;
        regexCheckbox.onClick = refreshPreview;

        prefixCheckbox.onClick = function () {
            updatePrefixEnabled();
            updateOkEnabled();
            rerender();
        };
        rbPrefixSepDash.onClick = rerender;
        rbPrefixSepUnderscore.onClick = rerender;
        prefixStartInput.onChanging = rerender;

        suffixCheckbox.onClick = function () {
            updateSuffixEnabled();
            updateOkEnabled();
            rerender();
        };
        rbSuffixSepDash.onClick = rerender;
        rbSuffixSepUnderscore.onClick = rerender;
        suffixStartInput.onChanging = rerender;

        updatePrefixEnabled();
        updateSuffixEnabled();
        findInput.active = true;
        refreshPreview();

        if (win.show() !== 1) {
            return;
        }

        var findText = findInput.text;
        var enablePrefix = prefixCheckbox.value;
        var enableSuffix = suffixCheckbox.value;

        if (findText === "" && !enablePrefix && !enableSuffix) {
            alert(L("needInput"));
            return;
        }

        var targetMode = getModeFromRadios();

        // リネーム計画 (実際に名前が変わる項目のみ) / Rename plan (only items whose name changes)
        var renamePlan = [];
        for (var rp = 0; rp < lastItems.length; rp++) {
            var entry = lastItems[rp];
            if (entry.newName !== entry.oldName) {
                renamePlan.push({
                    item: entry.item,
                    oldName: entry.oldName,
                    newName: entry.newName
                });
            }
        }

        /* リネーム計画から該当アイテムを検索 / Find planned rename for an item */
        function plannedRenameFor(refItem) {
            for (var i = 0; i < renamePlan.length; i++) {
                if (renamePlan[i].item === refItem) return renamePlan[i];
            }
            return null;
        }

        // =========================================
        // アートボード名 / Artboard names
        // =========================================

        /* アートボード名をリネーム / Rename artboards */
        function renameArtboards() {
            if (renamePlan.length === 0) {
                alert(L("noMatchArtboard"));
                return;
            }

            var count = 0;
            var errors = 0;

            for (var i = 0; i < renamePlan.length; i++) {
                try {
                    renamePlan[i].item.name = renamePlan[i].newName;
                    count++;
                } catch (e) {
                    errors++;
                }
            }

            alert(
                L("done") + "\n\n" +
                L("targetArtboards") + "\n" +
                L("renamed") + count + L("countSuffix") + "\n" +
                L("errorsLabel") + errors + L("countSuffix")
            );
        }

        // =========================================
        // レイヤー名 / Layer names
        // =========================================

        /* レイヤー名をリネーム / Rename layers */
        function renameLayers() {
            if (renamePlan.length === 0) {
                alert(L("noMatchLayer"));
                return;
            }

            var count = 0;
            var errors = 0;

            for (var i = 0; i < renamePlan.length; i++) {
                try {
                    renamePlan[i].item.name = renamePlan[i].newName;
                    count++;
                } catch (e) {
                    errors++;
                }
            }

            alert(
                L("done") + "\n\n" +
                L("targetLayers") + "\n" +
                L("renamed") + count + L("countSuffix") + "\n" +
                L("errorsLabel") + errors + L("countSuffix")
            );
        }

        // =========================================
        // シンボル名 / Symbol names
        // =========================================
        // シンボル名は同名不可のため、一度一時名にしてから最終名に変更します。
        // 同名になる場合は _2, _3 のように付けます。
        // Symbol names must be unique, so we rename via temporary names first.
        // Duplicates are suffixed as _2, _3, ...

        /* シンボル名をリネーム / Rename symbols */
        function renameSymbols() {
            if (renamePlan.length === 0) {
                alert(L("noMatchSymbol"));
                return;
            }

            var symbols = doc.symbols;
            var usedNames = {};

            // リネーム対象外のシンボル名を予約 / Reserve names of symbols not being renamed
            for (var i = 0; i < symbols.length; i++) {
                if (!plannedRenameFor(symbols[i])) {
                    usedNames[nameKey(symbols[i].name)] = true;
                }
            }

            var count = 0;
            var suffixed = 0;
            var errors = 0;

            // 一時名へ変更 / Rename to temporary names
            var tempPrefix = "__symbol_rename_temp__" + new Date().getTime() + "__";
            var renamedToTemp = [];

            for (var j = 0; j < renamePlan.length; j++) {
                try {
                    var tempName = tempPrefix + j;
                    renamePlan[j].item.name = tempName;
                    renamedToTemp.push(renamePlan[j]);
                } catch (e1) {
                    errors++;
                }
            }

            // 最終名へ変更 / Rename to final names
            for (var k = 0; k < renamedToTemp.length; k++) {
                try {
                    var desiredName = renamedToTemp[k].newName;
                    var finalName = makeUniqueName(desiredName, usedNames);

                    if (finalName !== desiredName) {
                        suffixed++;
                    }

                    renamedToTemp[k].item.name = finalName;
                    usedNames[nameKey(finalName)] = true;
                    count++;
                } catch (e2) {
                    errors++;
                }
            }

            alert(
                L("done") + "\n\n" +
                L("targetSymbols") + "\n" +
                L("renamed") + count + L("countSuffix") + "\n" +
                L("suffixed") + suffixed + L("countSuffix") + "\n" +
                L("errorsLabel") + errors + L("countSuffix")
            );
        }

        // =========================================
        // グラフィックスタイル名 / Graphic style names
        // =========================================
        // グラフィックスタイル名は同名不可のため、一度一時名にしてから最終名に変更します。
        // 初期スタイルなど、一部のスタイルは変更できない場合があります。
        // 同名になる場合は _2, _3 のように付けます。
        // Graphic style names must be unique, so we rename via temporary names first.
        // Some built-in styles may be immutable. Duplicates are suffixed as _2, _3, ...

        /* グラフィックスタイル名をリネーム / Rename graphic styles */
        function renameGraphicStyles() {
            if (renamePlan.length === 0) {
                alert(L("noMatchGraphicStyle"));
                return;
            }

            var styles = doc.graphicStyles;
            var usedNames = {};

            // リネーム対象外のスタイル名を予約 / Reserve names of styles not being renamed
            for (var i = 0; i < styles.length; i++) {
                if (!plannedRenameFor(styles[i])) {
                    try { usedNames[nameKey(styles[i].name)] = true; }
                    catch (e0) { /* 名前取得できないスタイルは無視 / Skip styles whose name is unavailable */ }
                }
            }

            var count = 0;
            var suffixed = 0;
            var errors = 0;

            // 一時名へ変更 / Rename to temporary names
            var tempPrefix = "__gstyle_rename_temp__" + new Date().getTime() + "__";
            var renamedToTemp = [];

            for (var j = 0; j < renamePlan.length; j++) {
                try {
                    var tempName = tempPrefix + j;
                    renamePlan[j].item.name = tempName;
                    renamedToTemp.push(renamePlan[j]);
                } catch (e1) {
                    errors++;
                }
            }

            // 最終名へ変更 / Rename to final names
            for (var k = 0; k < renamedToTemp.length; k++) {
                try {
                    var desiredName = renamedToTemp[k].newName;
                    var finalName = makeUniqueName(desiredName, usedNames);

                    if (finalName !== desiredName) {
                        suffixed++;
                    }

                    renamedToTemp[k].item.name = finalName;
                    usedNames[nameKey(finalName)] = true;
                    count++;
                } catch (e2) {
                    errors++;
                }
            }

            alert(
                L("done") + "\n\n" +
                L("targetGraphicStyles") + "\n" +
                L("renamed") + count + L("countSuffix") + "\n" +
                L("suffixed") + suffixed + L("countSuffix") + "\n" +
                L("errorsLabel") + errors + L("countSuffix")
            );
        }

        // =========================================
        // 実行 / Execute
        // =========================================

        if (targetMode === "artboard") {
            renameArtboards();
        } else if (targetMode === "layer") {
            renameLayers();
        } else if (targetMode === "symbol") {
            renameSymbols();
        } else if (targetMode === "graphicStyle") {
            renameGraphicStyles();
        }

    })();
