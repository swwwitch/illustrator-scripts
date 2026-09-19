#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

オブジェクトやレイヤーなどの名前を、条件を指定して一括で変更します。

詳細は README を参照してください。

### Overview

Renames objects, layers and the like in bulk, according to the conditions you set.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "renamer";                      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/renamer.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/renamer.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

    (function () {

    // =========================================
    // バージョンとローカライズ / Version and Localization
    // =========================================

    /* 現在のロケールを判定 / Detect current locale */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

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
        separator:           { ja: "区切り", en: "Separator" },
        startNumber:         { ja: "開始番号", en: "Start" },
        sort:                { ja: "並び替え", en: "Sort" },
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
        targetArtboards:     { ja: "対象：アートボード名", en: "Target: Artboards" },
        targetLayers:        { ja: "対象：レイヤー名", en: "Target: Layers" },
        targetSymbols:       { ja: "対象：シンボル名", en: "Target: Symbols" },
        targetGraphicStyles: { ja: "対象：グラフィックスタイル名", en: "Target: Graphic Styles" },
        renamed:             { ja: "変更数：", en: "Renamed: " },
        suffixed:            { ja: "同名回避で連番追加：", en: "Suffixed to avoid duplicates: " },
        errorsLabel:         { ja: "エラー：", en: "Errors: " },
        countSuffix:         { ja: " 件", en: "" },
        tipArtboard:         { ja: "アートボード名を対象にします。", en: "Renames artboard names." },
        tipLayer:            { ja: "レイヤー名を対象にします。", en: "Renames layer names." },
        tipSymbol:           { ja: "シンボル名を対象にします。", en: "Renames symbol names." },
        tipGraphicStyle:     { ja: "グラフィックスタイル名を対象にします。", en: "Renames graphic style names." },
        tipFindReplaceEnable:{ ja: "オフにすると、接頭辞・接尾辞のナンバリングだけを行います。", en: "When off, only the prefix/suffix numbering is applied." },
        tipFind:             { ja: "名前の中から探す文字列です。空欄だとすべてが対象になります。", en: "Text to look for in the name. Leave blank to target everything." },
        tipReplace:          { ja: "置き換える文字列です。空欄にすると検索文字列を削除します。", en: "Replacement text. Leave blank to delete the found text." },
        tipRegex:            { ja: "検索文字列を正規表現として扱います（$1 などの後方参照も使えます）。", en: "Treats the search text as a regular expression, including back-references such as $1." },
        tipPrefixEnable:     { ja: "名前の先頭に連番を付けます。", en: "Adds a sequential number to the start of the name." },
        tipSuffixEnable:     { ja: "名前の末尾に連番を付けます。", en: "Adds a sequential number to the end of the name." },
        tipSeparator:        { ja: "連番と名前の間に入れる文字です。", en: "Character placed between the number and the name." },
        tipStartNumber:      { ja: "連番の最初の数字です。", en: "First number of the sequence." },
        tipSort:             { ja: "プレビューと連番を振る順序です。", en: "Order used for the preview and for numbering." },
        tipMoveTop:          { ja: "選んだ項目を先頭へ移動します。", en: "Moves the selected item to the top." },
        tipMoveUp:           { ja: "選んだ項目を1つ上へ移動します。", en: "Moves the selected item up one position." },
        tipMoveDown:         { ja: "選んだ項目を1つ下へ移動します。", en: "Moves the selected item down one position." },
        tipMoveBottom:       { ja: "選んだ項目を末尾へ移動します。", en: "Moves the selected item to the bottom." },
        tipPreview:          { ja: "変更前と変更後の名前の一覧です。並べ替えたい項目はここで選びます。", en: "List of the names before and after the change. Select an item here to reorder it." }
    };

    /* ローカライズ文字列を取得 / Get localized string */
    function getLabel(key) {
        var entry = LABELS[key];
        if (!entry) return key;
        return entry[uiLang] || entry.en || entry.ja || key;
    }

    /* コロン付きの項目名を返す（日本語は全角、英語は半角） / Return a label with a colon */
    function labelText(key) {
        return getLabel(key) + (uiLang === "ja" ? "：" : ": ");
    }

        if (app.documents.length === 0) {
            alert(getLabel("noDoc"));
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

        var win = new Window("dialog", getLabel("dialogTitle") + " " + SCRIPT_VERSION);
        win.orientation = "column";
        win.alignChildren = ["fill", "top"];

        // 対象選択 (2カラムを貫通) / Target selection (spans both columns)
        var targetPanel = win.add("panel", undefined, getLabel("target"));
        targetPanel.orientation = "row";
        targetPanel.alignChildren = ["left", "center"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = 10;

        var rbArtboard = targetPanel.add("radiobutton", undefined, getLabel("artboard"));
        rbArtboard.helpTip = getLabel("tipArtboard");
        var rbLayer = targetPanel.add("radiobutton", undefined, getLabel("layer"));
        rbLayer.helpTip = getLabel("tipLayer");
        var rbSymbol = targetPanel.add("radiobutton", undefined, getLabel("symbol"));
        rbSymbol.helpTip = getLabel("tipSymbol");
        var rbGraphicStyle = targetPanel.add("radiobutton", undefined, getLabel("graphicStyle"));
        rbGraphicStyle.helpTip = getLabel("tipGraphicStyle");

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
        var findReplacePanel = leftColumn.add("panel", undefined, getLabel("findReplace"));
        setupPanel(findReplacePanel, 6);

        var findReplaceCheckbox = findReplacePanel.add("checkbox", undefined, getLabel("findReplaceEnable"));
        findReplaceCheckbox.helpTip = getLabel("tipFindReplaceEnable");
        findReplaceCheckbox.value = true;

        var findGroup = findReplacePanel.add("group");
        findGroup.orientation = "row";
        findGroup.alignChildren = ["left", "center"];
        findGroup.add("statictext", undefined, labelText("find"));
        var findInput = findGroup.add("edittext", undefined, "");
        findInput.helpTip = getLabel("tipFind");
        findInput.characters = 15;

        var replaceGroup = findReplacePanel.add("group");
        replaceGroup.orientation = "row";
        replaceGroup.alignChildren = ["left", "center"];
        replaceGroup.add("statictext", undefined, labelText("replace"));
        var replaceInput = replaceGroup.add("edittext", undefined, "");
        replaceInput.helpTip = getLabel("tipReplace");
        replaceInput.characters = 15;

        var regexGroup = findReplacePanel.add("group");
        regexGroup.orientation = "row";
        regexGroup.alignChildren = ["right", "center"];
        regexGroup.alignment = "fill";
        var regexCheckbox = regexGroup.add("checkbox", undefined, getLabel("regex"));
        regexCheckbox.helpTip = getLabel("tipRegex");

        // 接頭辞パネル / Prefix panel
        var prefixPanel = leftColumn.add("panel", undefined, getLabel("prefix"));
        setupPanel(prefixPanel, 6);

        var prefixCheckbox = prefixPanel.add("checkbox", undefined, getLabel("numberingEnable"));
        prefixCheckbox.helpTip = getLabel("tipPrefixEnable");

        var prefixSepGroup = prefixPanel.add("group");
        prefixSepGroup.orientation = "row";
        prefixSepGroup.alignChildren = ["left", "center"];
        prefixSepGroup.add("statictext", undefined, labelText("separator"));
        var rbPrefixSepDash = prefixSepGroup.add("radiobutton", undefined, "-");
        rbPrefixSepDash.helpTip = getLabel("tipSeparator");
        var rbPrefixSepUnderscore = prefixSepGroup.add("radiobutton", undefined, "_");
        rbPrefixSepUnderscore.helpTip = getLabel("tipSeparator");
        rbPrefixSepDash.value = true;

        var prefixStartGroup = prefixPanel.add("group");
        prefixStartGroup.orientation = "row";
        prefixStartGroup.alignChildren = ["left", "center"];
        prefixStartGroup.add("statictext", undefined, labelText("startNumber"));
        var prefixStartInput = prefixStartGroup.add("edittext", undefined, "1");
        prefixStartInput.helpTip = getLabel("tipStartNumber");
        prefixStartInput.characters = 4;

        // 接尾辞パネル / Suffix panel
        var suffixPanel = leftColumn.add("panel", undefined, getLabel("suffix"));
        setupPanel(suffixPanel, 6);

        var suffixCheckbox = suffixPanel.add("checkbox", undefined, getLabel("numberingEnable"));
        suffixCheckbox.helpTip = getLabel("tipSuffixEnable");

        var suffixSepGroup = suffixPanel.add("group");
        suffixSepGroup.orientation = "row";
        suffixSepGroup.alignChildren = ["left", "center"];
        suffixSepGroup.add("statictext", undefined, labelText("separator"));
        var rbSuffixSepDash = suffixSepGroup.add("radiobutton", undefined, "-");
        rbSuffixSepDash.helpTip = getLabel("tipSeparator");
        var rbSuffixSepUnderscore = suffixSepGroup.add("radiobutton", undefined, "_");
        rbSuffixSepUnderscore.helpTip = getLabel("tipSeparator");
        rbSuffixSepDash.value = true;

        var suffixStartGroup = suffixPanel.add("group");
        suffixStartGroup.orientation = "row";
        suffixStartGroup.alignChildren = ["left", "center"];
        suffixStartGroup.add("statictext", undefined, labelText("startNumber"));
        var suffixStartInput = suffixStartGroup.add("edittext", undefined, "1");
        suffixStartInput.helpTip = getLabel("tipStartNumber");
        suffixStartInput.characters = 4;

        // 右カラム: 並び替え + プレビュー / Right column: sort + preview
        var rightColumn = mainGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "fill"];

        var sortGroup = rightColumn.add("group");
        sortGroup.orientation = "row";
        sortGroup.alignChildren = ["left", "center"];
        sortGroup.add("statictext", undefined, labelText("sort"));
        var sortDropdown = sortGroup.add("dropdownlist", undefined, [
            getLabel("sortOriginal"),
            getLabel("sortNameAsc"),
            getLabel("sortNameDesc"),
            getLabel("sortChanged")
        ]);
        sortDropdown.selection = 0;
        sortDropdown.helpTip = getLabel("tipSort");

        var moveTopBtn = sortGroup.add("button", undefined, getLabel("moveTop"));
        moveTopBtn.helpTip = getLabel("tipMoveTop");
        var moveUpBtn = sortGroup.add("button", undefined, getLabel("moveUp"));
        moveUpBtn.helpTip = getLabel("tipMoveUp");
        var moveDownBtn = sortGroup.add("button", undefined, getLabel("moveDown"));
        moveDownBtn.helpTip = getLabel("tipMoveDown");
        var moveBottomBtn = sortGroup.add("button", undefined, getLabel("moveBottom"));
        moveBottomBtn.helpTip = getLabel("tipMoveBottom");
        moveTopBtn.preferredSize.width = 36;
        moveUpBtn.preferredSize.width = 32;
        moveDownBtn.preferredSize.width = 32;
        moveBottomBtn.preferredSize.width = 36;

        var previewList = rightColumn.add("listbox", undefined, []);
        previewList.helpTip = getLabel("tipPreview");
        previewList.preferredSize.width = 340;
        previewList.preferredSize.height = PREVIEW_VISIBLE_LINES * PREVIEW_LINE_HEIGHT;

        // ボタン (Mac規約: Cancel → OK) / Buttons (Mac convention)
        var buttonGroup = win.add("group");
        buttonGroup.alignment = "right";

        buttonGroup.add("button", undefined, getLabel("cancel"), { name: "cancel" });
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
            var currentSelection = previewList.selection;
            if (!currentSelection) return;
            var idx = currentSelection.index;
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
            alert(getLabel("needInput"));
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
                alert(getLabel("noMatchArtboard"));
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
                getLabel("done") + "\n\n" +
                getLabel("targetArtboards") + "\n" +
                getLabel("renamed") + count + getLabel("countSuffix") + "\n" +
                getLabel("errorsLabel") + errors + getLabel("countSuffix")
            );
        }

        // =========================================
        // レイヤー名 / Layer names
        // =========================================

        /* レイヤー名をリネーム / Rename layers */
        function renameLayers() {
            if (renamePlan.length === 0) {
                alert(getLabel("noMatchLayer"));
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
                getLabel("done") + "\n\n" +
                getLabel("targetLayers") + "\n" +
                getLabel("renamed") + count + getLabel("countSuffix") + "\n" +
                getLabel("errorsLabel") + errors + getLabel("countSuffix")
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
                alert(getLabel("noMatchSymbol"));
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
                getLabel("done") + "\n\n" +
                getLabel("targetSymbols") + "\n" +
                getLabel("renamed") + count + getLabel("countSuffix") + "\n" +
                getLabel("suffixed") + suffixed + getLabel("countSuffix") + "\n" +
                getLabel("errorsLabel") + errors + getLabel("countSuffix")
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
                alert(getLabel("noMatchGraphicStyle"));
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
                getLabel("done") + "\n\n" +
                getLabel("targetGraphicStyles") + "\n" +
                getLabel("renamed") + count + getLabel("countSuffix") + "\n" +
                getLabel("suffixed") + suffixed + getLabel("countSuffix") + "\n" +
                getLabel("errorsLabel") + errors + getLabel("countSuffix")
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
