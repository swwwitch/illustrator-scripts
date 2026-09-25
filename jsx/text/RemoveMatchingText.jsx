#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

入力した文字列（5つまで、正規表現も可）を、選択中のオブジェクト・現在のアートボード・ドキュメント全体のテキストから削除します。
残った文字の書式は変わりません。シンボル内のテキストや、非表示・ロックされたテキストも対象にできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RemoveMatchingText.md

### Overview

Removes up to five strings (regular expressions allowed) from text in the selection, the current artboard, or the entire document.
The formatting of the remaining text is kept. Text in symbols and hidden or locked text can be included.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RemoveMatchingText.md

*/

(function () {

    // =========================================
    // 基本情報 / Basic info
    // =========================================
    var SCRIPT_NAME     = "RemoveMatchingText";           /* スクリプト名 / script name */
    var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
    var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
    var SCRIPT_RELEASED = "2026-09-26";                   /* 最初のリリース日 / first release date */
    var SCRIPT_UPDATED  = "2026-09-26";                   /* 更新日 / last updated */

    var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RemoveMatchingText.md"; /* README（日本語） */
    var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RemoveMatchingText.md"; /* README (English) */

    // Released under the MIT license
    // http://opensource.org/licenses/mit-license.php

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var SEARCH_FIELD_COUNT = 5;             /* 削除する文字列の入力欄の数 / number of search fields */
    var DEFAULT_USE_REGEX = false;          /* 正規表現（初期値）/ regular expression (initial value) */
    var DEFAULT_IGNORE_CASE = false;        /* 大文字と小文字を区別しない（初期値）/ ignore case (initial value) */
    var DEFAULT_DELETE_EMPTY_FRAMES = true; /* 空になったフレームを削除（初期値）/ delete emptied frames (initial value) */
    var DEFAULT_INCLUDE_HIDDEN = false;     /* 非表示のレイヤーを検索（初期値）/ search hidden layers (initial value) */
    var DEFAULT_INCLUDE_LOCKED = false;     /* ロックされたレイヤーを検索（初期値）/ search locked layers (initial value) */
    var DEFAULT_INCLUDE_SYMBOLS = true;     /* シンボルも検索（初期値）/ search symbols too (initial value) */

    // =========================================
    // 前回の設定 / Saved settings
    // =========================================
    var SETTINGS_PREF_KEY = "RemoveMatchingText/settings"; /* 環境設定に保存するキー / preference key */

    /**
     * 前回 OK したときの設定を読み込む
     * @returns {Object|null} 保存した設定。無いか壊れていれば null
     */
    function loadSettings() {
        var savedText = app.preferences.getStringPreference(SETTINGS_PREF_KEY);
        if (!savedText) return null;
        /* 壊れた文字列は eval が例外を出す / eval throws on a corrupted string */
        try {
            return eval(savedText);
        } catch (e) {
            return null;
        }
    }

    /**
     * 設定を環境設定に保存する（Illustrator を再起動しても残る）
     * @param {Object} settings - 保存する設定
     * @returns {void}
     */
    function saveSettings(settings) {
        app.preferences.setStringPreference(SETTINGS_PREF_KEY, settings.toSource());
    }

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS = [15, 20, 15, 10];   /* パネルの内側余白 / panel margins */
    var PANEL_SPACING = 6;                  /* パネル内の間隔 / panel spacing */
    var INPUT_CHARACTERS = 20;              /* 入力欄の幅（文字数）/ input width in characters */
    var MATCH_COUNT_WIDTH = 40;             /* 一致数の表示幅 / width of the match count */
    var SEARCH_OPTIONS_TOP_MARGIN = 5;      /* 正規表現などの上余白 / top margin above the search options */
    var BUTTON_ROW_TOP_MARGIN = 6;          /* ボタン行の上余白 / top margin of the button row */

    /* パネルの共通設定 / Shared panel setup */
    function setupPanel(targetPanel) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["left", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = PANEL_SPACING;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "指定した文字列を削除", en: "Remove Specified Text" }
        },
        panel: {
            searchText: { ja: "削除する文字列", en: "Text to Remove" },
            scope: { ja: "対象", en: "Scope" },
            options: { ja: "オプション", en: "Options" }
        },
        checkbox: {
            useRegex: { ja: "正規表現", en: "Regular expression" },
            ignoreCase: { ja: "大文字と小文字を区別しない", en: "Ignore case" },
            deleteEmptyFrames: { ja: "空になったテキストを削除", en: "Delete emptied text" },
            includeHidden: { ja: "非表示のレイヤーを検索", en: "Search hidden layers" },
            includeLocked: { ja: "ロックされたレイヤーを検索", en: "Search locked layers" },
            includeSymbols: { ja: "シンボルも検索", en: "Search symbols too" }
        },
        radio: {
            selection: { ja: "選択中のオブジェクト", en: "Selected objects" },
            artboard: { ja: "現在のアートボード", en: "Current artboard" },
            wholeDocument: { ja: "ドキュメント全体", en: "Entire document" }
        },
        tooltip: {
            searchText: {
                ja: "一致する部分をすべて削除します。空欄は無視し、上の欄から順に処理します。入力内容は次回に引き継がれます",
                en: "Removes every match. Empty fields are ignored; fields are processed from top to bottom. Entries are kept for the next run"
            },
            matchCount: {
                ja: "対象の範囲で見つかった数（ほかの欄による削除は考慮しない）。「!」は正規表現の誤り",
                en: "Matches found in the scope (ignoring removals by other fields). \"!\" means an invalid regular expression"
            },
            useRegex: {
                ja: "入力を JavaScript の正規表現として扱います。^ と $ は段落の先頭・末尾に一致します",
                en: "Treats the input as JavaScript regular expressions. ^ and $ match the start and end of each paragraph"
            },
            selection: {
                ja: "グループ内のテキストも対象になります",
                en: "Includes text inside groups"
            },
            artboard: {
                ja: "アートボードに一部でも重なるテキストが対象になります",
                en: "Includes text that partly overlaps the artboard"
            },
            deleteEmptyFrames: {
                ja: "今回の削除で空になったフレームだけを削除します。元から空のフレームと、スレッドテキスト（連結）のフレームは残します",
                en: "Deletes only frames emptied by this run. Frames that were already empty and threaded text frames are kept"
            },
            includeHidden: {
                ja: "非表示のレイヤー・オブジェクト内のテキストも対象にします。処理のあいだだけ表示し、終わったら元に戻します",
                en: "Includes text in hidden layers and objects. They are shown only while processing, then hidden again"
            },
            includeLocked: {
                ja: "ロックされたレイヤー・オブジェクト内のテキストも対象にします。処理のあいだだけロックを解除し、終わったら元に戻します",
                en: "Includes text in locked layers and objects. They are unlocked only while processing, then locked again"
            },
            includeSymbols: {
                ja: "対象範囲にあるシンボルの定義を書き換えます。範囲外にある同じシンボルのインスタンスも変わり、基準点などのシンボルオプションは初期値になります",
                en: "Rewrites the definitions of symbols in the scope. Instances outside the scope change too, and symbol options such as the registration point are reset"
            },
            selectionUnavailable: {
                ja: "オブジェクトが選択されていないため選べません",
                en: "Unavailable because nothing is selected"
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            removedCounts: { ja: "削除した数", en: "Removed matches" },
            removedCountLine: { ja: "「{text}」：{count}", en: "\"{text}\": {count}" },
            result: {
                ja: "{count}個のテキストオブジェクトから削除しました。",
                en: "Removed the text from {count} text object(s)."
            },
            deletedFrames: {
                ja: "空になった{count}個のテキストフレームを削除しました。",
                en: "Deleted {count} text frame(s) left empty."
            },
            updatedSymbols: { ja: "{count}個のシンボルを書き換えました。", en: "Rewrote {count} symbol(s)." }
        }
    };

    /**
     * 現在の言語のラベルを返す
     * @param {Object} labelSet - { ja, en } を持つラベル
     * @returns {string} ラベル文字列
     */
    function getLabel(labelSet) {
        return labelSet[uiLang] || labelSet.en;
    }

    /**
     * コロン付きのラベルを返す（日本語は全角、英語は半角）
     * @param {Object} labelSet - { ja, en } を持つラベル
     * @returns {string} コロン付きラベル
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * メイン処理
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        var doc = app.activeDocument;
        /* シンボルのリンク解除で選択が変わるので、最初に控えておく / Keep the selection now, as breaking symbol links changes it */
        var selectedItems = toItemArray(doc.selection);

        var removeOptions = showRemoveDialog(doc, selectedItems);
        if (!removeOptions) return;

        var targetFrames = collectTargetFrames(doc, removeOptions.scope, removeOptions.layerOptions, selectedItems);
        var removeResult = removeTextFromFrames(targetFrames, removeOptions.searchEntries, removeOptions.deleteEmptyFrames);
        removeResult.updatedSymbolCount = 0;
        if (removeOptions.includeSymbols) {
            var targetSymbols = collectTargetSymbols(doc, removeOptions.scope, removeOptions.layerOptions, selectedItems);
            removeResult.updatedSymbolCount = removeTextFromSymbols(doc, targetSymbols, removeOptions.searchEntries, removeResult.removedCounts);
        }
        alert(buildResultMessage(removeOptions.searchEntries, removeResult));
    }

    /**
     * 選択（配列でない TextRange などを含む）をアイテムの配列にする
     * @param {Object} selection - doc.selection
     * @returns {PageItem[]} アイテムの配列。選択が無ければ空
     */
    function toItemArray(selection) {
        var items = [];
        if (!selection || !selection.length) return items;
        for (var i = 0; i < selection.length; i++) {
            items.push(selection[i]);
        }
        return items;
    }

    /**
     * 結果メッセージを組み立てる（欄ごとの削除数・変更したオブジェクト数・削除したフレーム数）
     * @param {Object[]} searchEntries - getSearchEntries() の結果
     * @param {Object} removeResult - removeTextFromFrames() の結果
     * @returns {string} メッセージ
     */
    function buildResultMessage(searchEntries, removeResult) {
        var messageLines = [labelText(LABELS.alert.removedCounts)];
        for (var i = 0; i < searchEntries.length; i++) {
            messageLines.push(getLabel(LABELS.alert.removedCountLine)
                .replace("{text}", searchEntries[i].text)
                .replace("{count}", removeResult.removedCounts[i]));
        }
        messageLines.push("");
        messageLines.push(getLabel(LABELS.alert.result).replace("{count}", removeResult.changedCount));
        if (removeResult.deletedFrameCount > 0) {
            messageLines.push(getLabel(LABELS.alert.deletedFrames).replace("{count}", removeResult.deletedFrameCount));
        }
        if (removeResult.updatedSymbolCount > 0) {
            messageLines.push(getLabel(LABELS.alert.updatedSymbols).replace("{count}", removeResult.updatedSymbolCount));
        }
        return messageLines.join("\n");
    }

    /**
     * ダイアログを表示し、削除するパターンと対象範囲を返す
     * @param {Document} doc - 一致数を数えるドキュメント
     * @param {PageItem[]} selectedItems - 実行時に選択していたアイテム
     * @returns {Object|null} { searchEntries: Object[], scope: string, deleteEmptyFrames: boolean, includeSymbols: boolean, layerOptions: Object }。キャンセル時は null
     */
    function showRemoveDialog(doc, selectedItems) {
        /* 前回の設定があれば初期値にする / Start from the saved settings when present */
        var savedSettings = loadSettings() || {};
        var savedTexts = savedSettings.searchTexts || [];
        function savedValue(key, defaultValue) {
            return (typeof savedSettings[key] === "boolean") ? savedSettings[key] : defaultValue;
        }

        var dlg = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];

        /* 削除する文字列 / Text to remove */
        var searchPanel = dlg.add("panel", undefined, getLabel(LABELS.panel.searchText));
        setupPanel(searchPanel);
        var searchInputs = [];
        var matchCountLabels = [];
        for (var i = 0; i < SEARCH_FIELD_COUNT; i++) {
            var searchRowGroup = searchPanel.add("group");
            searchRowGroup.orientation = "row";
            searchRowGroup.alignChildren = ["left", "center"];
            var searchInput = searchRowGroup.add("edittext", undefined, savedTexts[i] || "");
            searchInput.characters = INPUT_CHARACTERS;
            searchInput.helpTip = getLabel(LABELS.tooltip.searchText);
            searchInput.onChanging = refreshDialogState;
            var matchCountLabel = searchRowGroup.add("statictext", undefined, "");
            matchCountLabel.preferredSize.width = MATCH_COUNT_WIDTH;
            matchCountLabel.justify = "right";
            matchCountLabel.helpTip = getLabel(LABELS.tooltip.matchCount);
            searchInputs.push(searchInput);
            matchCountLabels.push(matchCountLabel);
        }
        searchInputs[0].active = true;
        var searchOptionsGroup = searchPanel.add("group");
        searchOptionsGroup.orientation = "column";
        searchOptionsGroup.alignChildren = ["left", "top"];
        searchOptionsGroup.margins = [0, SEARCH_OPTIONS_TOP_MARGIN, 0, 0];
        searchOptionsGroup.spacing = PANEL_SPACING;
        var useRegexCheckbox = searchOptionsGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.useRegex));
        useRegexCheckbox.value = savedValue("useRegex", DEFAULT_USE_REGEX);
        useRegexCheckbox.helpTip = getLabel(LABELS.tooltip.useRegex);
        useRegexCheckbox.onClick = refreshDialogState;
        var ignoreCaseCheckbox = searchOptionsGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.ignoreCase));
        ignoreCaseCheckbox.value = savedValue("ignoreCase", DEFAULT_IGNORE_CASE);
        ignoreCaseCheckbox.onClick = refreshDialogState;

        /* 対象 / Scope */
        var scopePanel = dlg.add("panel", undefined, getLabel(LABELS.panel.scope));
        setupPanel(scopePanel);
        var selectionRadio = scopePanel.add("radiobutton", undefined, getLabel(LABELS.radio.selection));
        selectionRadio.helpTip = getLabel(LABELS.tooltip.selection);
        var artboardRadio = scopePanel.add("radiobutton", undefined, getLabel(LABELS.radio.artboard));
        artboardRadio.helpTip = getLabel(LABELS.tooltip.artboard);
        var documentRadio = scopePanel.add("radiobutton", undefined, getLabel(LABELS.radio.wholeDocument));
        /* 何も選択していなければ「選択中のオブジェクト」は選べず、ドキュメント全体にする / Without a selection, fall back to the whole document */
        var hasSelection = selectedItems.length > 0;
        if (!hasSelection) {
            selectionRadio.enabled = false;
            selectionRadio.helpTip = getLabel(LABELS.tooltip.selectionUnavailable);
        }
        if (savedSettings.scope === "artboard") artboardRadio.value = true;
        else if (savedSettings.scope === "document" || !hasSelection) documentRadio.value = true;
        else selectionRadio.value = true;
        selectionRadio.onClick = artboardRadio.onClick = documentRadio.onClick = refreshDialogState;

        /* オプション / Options */
        var optionsPanel = dlg.add("panel", undefined, getLabel(LABELS.panel.options));
        setupPanel(optionsPanel);
        var deleteEmptyFramesCheckbox = optionsPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.deleteEmptyFrames));
        deleteEmptyFramesCheckbox.value = savedValue("deleteEmptyFrames", DEFAULT_DELETE_EMPTY_FRAMES);
        deleteEmptyFramesCheckbox.helpTip = getLabel(LABELS.tooltip.deleteEmptyFrames);
        var includeHiddenCheckbox = optionsPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.includeHidden));
        includeHiddenCheckbox.value = savedValue("includeHidden", DEFAULT_INCLUDE_HIDDEN);
        includeHiddenCheckbox.helpTip = getLabel(LABELS.tooltip.includeHidden);
        includeHiddenCheckbox.onClick = refreshDialogState;
        var includeLockedCheckbox = optionsPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.includeLocked));
        includeLockedCheckbox.value = savedValue("includeLocked", DEFAULT_INCLUDE_LOCKED);
        includeLockedCheckbox.helpTip = getLabel(LABELS.tooltip.includeLocked);
        includeLockedCheckbox.onClick = refreshDialogState;
        var includeSymbolsCheckbox = optionsPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.includeSymbols));
        includeSymbolsCheckbox.value = savedValue("includeSymbols", DEFAULT_INCLUDE_SYMBOLS);
        includeSymbolsCheckbox.helpTip = getLabel(LABELS.tooltip.includeSymbols);
        includeSymbolsCheckbox.onClick = refreshDialogState;

        /* ボタン / Buttons */
        var btnRowGroup = dlg.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        /* 対象範囲ごとのテキスト内容（範囲やレイヤーの扱いを切り替えたときだけ集め直す）/ Contents cached per scope and layer options */
        var contentsCache = {};

        /* 選択中の対象範囲 / Currently selected scope */
        function getSelectedScope() {
            if (selectionRadio.value) return "selection";
            if (artboardRadio.value) return "artboard";
            return "document";
        }

        /* 非表示・ロックの扱い / How hidden and locked layers are handled */
        function getLayerOptions() {
            return { includeHidden: includeHiddenCheckbox.value, includeLocked: includeLockedCheckbox.value };
        }

        /* 一致数と OK の可否を更新（有効なパターンが無いか、誤った正規表現がある間は OK を押せない）
           Update match counts; disable OK when no pattern is given or a regular expression is invalid */
        function refreshDialogState() {
            var scope = getSelectedScope();
            var layerOptions = getLayerOptions();
            var includeSymbols = includeSymbolsCheckbox.value;
            var cacheKey = [scope, layerOptions.includeHidden, layerOptions.includeLocked, includeSymbols].join(":");
            if (!contentsCache[cacheKey]) {
                var textContents = getFrameContents(collectTargetFrames(doc, scope, layerOptions, selectedItems));
                if (includeSymbols) {
                    textContents = textContents.concat(getSymbolTextContents(doc, collectTargetSymbols(doc, scope, layerOptions, selectedItems)));
                }
                contentsCache[cacheKey] = textContents;
            }
            var hasPattern = false;
            var hasInvalidPattern = false;
            for (var i = 0; i < searchInputs.length; i++) {
                var searchPattern = createSearchPattern(searchInputs[i].text, useRegexCheckbox.value, ignoreCaseCheckbox.value);
                if (searchPattern === null) {
                    matchCountLabels[i].text = "";
                } else if (searchPattern === false) {
                    matchCountLabels[i].text = "!";
                    hasInvalidPattern = true;
                } else {
                    matchCountLabels[i].text = String(countMatches(contentsCache[cacheKey], searchPattern));
                    hasPattern = true;
                }
            }
            btnOK.enabled = hasPattern && !hasInvalidPattern;
        }

        refreshDialogState();

        if (dlg.show() !== 1) return null;

        var layerOptions = getLayerOptions();
        var searchTexts = [];
        for (var j = 0; j < searchInputs.length; j++) {
            searchTexts.push(searchInputs[j].text);
        }
        saveSettings({
            searchTexts: searchTexts,
            useRegex: useRegexCheckbox.value,
            ignoreCase: ignoreCaseCheckbox.value,
            scope: getSelectedScope(),
            deleteEmptyFrames: deleteEmptyFramesCheckbox.value,
            includeHidden: layerOptions.includeHidden,
            includeLocked: layerOptions.includeLocked,
            includeSymbols: includeSymbolsCheckbox.value
        });

        return {
            searchEntries: getSearchEntries(searchTexts, useRegexCheckbox.value, ignoreCaseCheckbox.value),
            scope: getSelectedScope(),
            deleteEmptyFrames: deleteEmptyFramesCheckbox.value,
            includeSymbols: includeSymbolsCheckbox.value,
            layerOptions: layerOptions
        };
    }

    /**
     * 入力から検索パターンを作る
     * @param {string} searchText - 入力された文字列
     * @param {boolean} useRegex - 正規表現として扱うか
     * @param {boolean} ignoreCase - 大文字と小文字を区別しないか
     * @returns {RegExp|null|boolean} パターン。空欄なら null、正規表現が誤っていれば false
     */
    function createSearchPattern(searchText, useRegex, ignoreCase) {
        if (searchText === "") return null;
        var source = useRegex ? searchText : searchText.replace(/[.*+?^${}()|[\]\\\/]/g, "\\$&");
        /* 誤った正規表現は new RegExp() が例外を出す / new RegExp() throws on an invalid pattern */
        try {
            return new RegExp(source, ignoreCase ? "gmi" : "gm");
        } catch (e) {
            return false;
        }
    }

    /**
     * 入力から有効な検索パターンを上から順に集める
     * @param {string[]} searchTexts - 入力欄の文字列
     * @param {boolean} useRegex - 正規表現として扱うか
     * @param {boolean} ignoreCase - 大文字と小文字を区別しないか
     * @returns {Object[]} { text: string, pattern: RegExp } の配列
     */
    function getSearchEntries(searchTexts, useRegex, ignoreCase) {
        var searchEntries = [];
        for (var i = 0; i < searchTexts.length; i++) {
            var searchPattern = createSearchPattern(searchTexts[i], useRegex, ignoreCase);
            if (searchPattern) searchEntries.push({ text: searchTexts[i], pattern: searchPattern });
        }
        return searchEntries;
    }

    /**
     * テキストフレームの内容を配列にする
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {string[]} 各フレームの contents
     */
    function getFrameContents(textFrames) {
        var frameContents = [];
        for (var i = 0; i < textFrames.length; i++) {
            frameContents.push(textFrames[i].contents);
        }
        return frameContents;
    }

    /**
     * 文字列の配列に含まれる一致数の合計を返す
     * @param {string[]} frameContents - 各フレームの contents
     * @param {RegExp} searchPattern - 検索パターン
     * @returns {number} 一致数の合計
     */
    function countMatches(frameContents, searchPattern) {
        var matchCount = 0;
        for (var i = 0; i < frameContents.length; i++) {
            matchCount += findMatches(frameContents[i], searchPattern).length;
        }
        return matchCount;
    }

    /**
     * 対象範囲のテキストフレームを集める
     * @param {Document} doc - 対象ドキュメント
     * @param {string} scope - "selection" / "artboard" / "document"
     * @param {Object} layerOptions - { includeHidden: boolean, includeLocked: boolean }
     * @param {PageItem[]} selectedItems - 実行時に選択していたアイテム
     * @returns {TextFrame[]} テキストフレームの配列
     */
    function collectTargetFrames(doc, scope, layerOptions, selectedItems) {
        return collectScopedItems(doc.textFrames, "TextFrame", doc, scope, layerOptions, selectedItems);
    }

    /**
     * 対象範囲にあるアイテムを、指定した型だけ集める
     * @param {Object} documentItems - ドキュメント全体のコレクション（doc.textFrames / doc.symbolItems）
     * @param {string} typename - 集める型名
     * @param {Document} doc - 対象ドキュメント
     * @param {string} scope - "selection" / "artboard" / "document"
     * @param {Object} layerOptions - { includeHidden: boolean, includeLocked: boolean }
     * @param {PageItem[]} selectedItems - 実行時に選択していたアイテム
     * @returns {PageItem[]} アイテムの配列
     */
    function collectScopedItems(documentItems, typename, doc, scope, layerOptions, selectedItems) {
        var candidateItems = [];
        var i;

        if (scope === "selection") {
            for (i = 0; i < selectedItems.length; i++) {
                collectItemsOfType(selectedItems[i], typename, candidateItems);
            }
        } else {
            var artboardRect = null;
            if (scope === "artboard") {
                artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            }
            for (i = 0; i < documentItems.length; i++) {
                if (!artboardRect || overlapsArtboard(documentItems[i], artboardRect)) {
                    candidateItems.push(documentItems[i]);
                }
            }
        }

        var scopedItems = [];
        for (i = 0; i < candidateItems.length; i++) {
            if (isSearchableItem(candidateItems[i], layerOptions)) scopedItems.push(candidateItems[i]);
        }
        return scopedItems;
    }

    /**
     * 非表示・ロックの扱いに照らして、検索対象にできるか判定する（自身と親のグループ・レイヤーを見る）
     * @param {PageItem} item - 判定するアイテム
     * @param {Object} layerOptions - { includeHidden: boolean, includeLocked: boolean }
     * @returns {boolean} 対象にできれば true
     */
    function isSearchableItem(item, layerOptions) {
        for (var node = item; node && node.typename !== "Document"; node = node.parent) {
            var isHidden = (node.typename === "Layer") ? !node.visible : node.hidden;
            if (isHidden && !layerOptions.includeHidden) return false;
            if (node.locked && !layerOptions.includeLocked) return false;
        }
        return true;
    }

    /**
     * 自身と親の非表示・ロックを一時的に解除し、元に戻すための記録を追加する
     * @param {PageItem} item - 対象のアイテム
     * @param {Object[]} stateRecords - { target, propertyName, originalValue } の追加先
     * @returns {void}
     */
    function unlockAndRevealAncestors(item, stateRecords) {
        for (var node = item; node && node.typename !== "Document"; node = node.parent) {
            /* ロック中は表示を切り替えられないことがあるので、先にロックを外す / Unlock first, as a locked item may refuse to change visibility */
            if (node.locked) {
                stateRecords.push({ target: node, propertyName: "locked", originalValue: true });
                node.locked = false;
            }
            if (node.typename === "Layer") {
                if (!node.visible) {
                    stateRecords.push({ target: node, propertyName: "visible", originalValue: false });
                    node.visible = true;
                }
            } else if (node.hidden) {
                stateRecords.push({ target: node, propertyName: "hidden", originalValue: true });
                node.hidden = false;
            }
        }
    }

    /**
     * 一時的に変えた非表示・ロックを元に戻す（変えたときと逆の順に）
     * @param {Object[]} stateRecords - unlockAndRevealAncestors() の記録
     * @returns {void}
     */
    function restoreItemStates(stateRecords) {
        for (var i = stateRecords.length - 1; i >= 0; i--) {
            stateRecords[i].target[stateRecords[i].propertyName] = stateRecords[i].originalValue;
        }
    }

    /**
     * アイテム内から指定した型のアイテムを再帰的に集める（グループ内も含む）
     * @param {PageItem} item - 調べるアイテム
     * @param {string} typename - 集める型名
     * @param {PageItem[]} result - 見つかったアイテムの追加先
     * @returns {void}
     */
    function collectItemsOfType(item, typename, result) {
        if (item.typename === typename) {
            result.push(item);
        } else if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) {
                collectItemsOfType(item.pageItems[i], typename, result);
            }
        }
    }

    /**
     * アイテムがアートボードに重なっているか判定する
     * @param {PageItem} item - 判定するアイテム
     * @param {number[]} artboardRect - アートボードの [左, 上, 右, 下]
     * @returns {boolean} 一部でも重なっていれば true
     */
    function overlapsArtboard(item, artboardRect) {
        var bounds = item.visibleBounds;
        return bounds[2] > artboardRect[0] &&
            bounds[0] < artboardRect[2] &&
            bounds[3] < artboardRect[1] &&
            bounds[1] > artboardRect[3];
    }

    /**
     * テキストフレームから一致箇所をすべて削除する（パターンは配列の順に処理）
     * @param {TextFrame[]} targetFrames - 対象のテキストフレーム
     * @param {Object[]} searchEntries - getSearchEntries() の結果
     * @param {boolean} deleteEmptyFrames - 削除で空になった独立フレームを削除するか
     * @returns {Object} { removedCounts: number[]（欄ごとの削除数）, changedCount: number, deletedFrameCount: number }
     */
    function removeTextFromFrames(targetFrames, searchEntries, deleteEmptyFrames) {
        var removeResult = { removedCounts: [], changedCount: 0, deletedFrameCount: 0 };
        for (var k = 0; k < searchEntries.length; k++) {
            removeResult.removedCounts.push(0);
        }
        for (var i = 0; i < targetFrames.length; i++) {
            var textFrame = targetFrames[i];
            var stateRecords = [];
            /* 編集できないテキスト（テンプレートレイヤーなど）は例外になるのでスキップ / Skip text that throws because it cannot be edited (template layers, etc.) */
            try {
                unlockAndRevealAncestors(textFrame, stateRecords);
                var isChanged = removePatternsFromFrame(textFrame, searchEntries, removeResult.removedCounts);
                if (isChanged) removeResult.changedCount++;
                if (isChanged && deleteEmptyFrames && textFrame.contents === "" && isStandaloneFrame(textFrame)) {
                    /* 消したフレーム自身は状態を戻さない / Do not restore the deleted frame itself */
                    stateRecords = excludeRecordsFor(stateRecords, textFrame);
                    textFrame.remove();
                    removeResult.deletedFrameCount++;
                }
            } catch (e) {
            } finally {
                restoreItemStates(stateRecords);
            }
        }
        return removeResult;
    }

    /**
     * 1つのテキストフレームから、パターンを配列の順にすべて削除する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object[]} searchEntries - getSearchEntries() の結果
     * @param {number[]} removedCounts - 欄ごとの削除数（加算していく）
     * @returns {boolean} 1か所でも削除したら true
     */
    function removePatternsFromFrame(textFrame, searchEntries, removedCounts) {
        var isChanged = false;
        for (var i = 0; i < searchEntries.length; i++) {
            /* 前の削除で内容が変わるので、毎回 contents から探し直す / Search again each time, as earlier removals change the contents */
            var matches = findMatches(textFrame.contents, searchEntries[i].pattern);
            if (matches.length === 0) continue;
            removeMatchesKeepingFormat(textFrame, matches);
            removedCounts[i] += matches.length;
            isChanged = true;
        }
        return isChanged;
    }

    /**
     * スレッドテキスト（連結）に属さない独立したフレームか判定する
     * @param {TextFrame} textFrame - 判定するテキストフレーム
     * @returns {boolean} 独立したフレームなら true
     */
    function isStandaloneFrame(textFrame) {
        if (textFrame.kind === TextType.POINTTEXT) return true;
        return !textFrame.nextFrame && !textFrame.previousFrame;
    }

    /**
     * 指定したオブジェクトの記録を除いた配列を返す
     * @param {Object[]} stateRecords - unlockAndRevealAncestors() の記録
     * @param {PageItem} target - 除くオブジェクト
     * @returns {Object[]} 残した記録
     */
    function excludeRecordsFor(stateRecords, target) {
        var keptRecords = [];
        for (var i = 0; i < stateRecords.length; i++) {
            if (stateRecords[i].target !== target) keptRecords.push(stateRecords[i]);
        }
        return keptRecords;
    }

    /**
     * 文字列中の一致箇所を左から集める（長さ0の一致は除く）
     * @param {string} text - 検索する文字列
     * @param {RegExp} searchPattern - g フラグ付きの検索パターン
     * @returns {Object[]} { start: number, length: number } の配列（昇順）
     */
    function findMatches(text, searchPattern) {
        var matches = [];
        var match;
        searchPattern.lastIndex = 0;
        while ((match = searchPattern.exec(text)) !== null) {
            if (match[0].length === 0) {
                /* 長さ0の一致で止まらないよう1文字進める / Step past empty matches to avoid an endless loop */
                searchPattern.lastIndex++;
                continue;
            }
            matches.push({ start: match.index, length: match[0].length });
        }
        return matches;
    }

    /**
     * 一致箇所を1文字ずつ右から削除する（contents を書き戻さないので文字ごとの書式が残る）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object[]} matches - findMatches() の結果（昇順）
     * @returns {void}
     */
    function removeMatchesKeepingFormat(textFrame, matches) {
        var frameCharacters = textFrame.characters;
        for (var i = matches.length - 1; i >= 0; i--) {
            for (var j = matches[i].start + matches[i].length - 1; j >= matches[i].start; j--) {
                frameCharacters[j].remove();
            }
        }
    }

    // =========================================
    // シンボル内のテキスト / Text in symbols
    // =========================================

    /**
     * 対象範囲にあるシンボルインスタンスから、元のシンボルを重複なく集める
     * @param {Document} doc - 対象ドキュメント
     * @param {string} scope - "selection" / "artboard" / "document"
     * @param {Object} layerOptions - { includeHidden: boolean, includeLocked: boolean }
     * @param {PageItem[]} selectedItems - 実行時に選択していたアイテム
     * @returns {Symbol[]} シンボルの配列
     */
    function collectTargetSymbols(doc, scope, layerOptions, selectedItems) {
        var symbolItems = collectScopedItems(doc.symbolItems, "SymbolItem", doc, scope, layerOptions, selectedItems);
        var targetSymbols = [];
        for (var i = 0; i < symbolItems.length; i++) {
            var symbol = symbolItems[i].symbol;
            if (indexOfItem(targetSymbols, symbol) === -1) targetSymbols.push(symbol);
        }
        return targetSymbols;
    }

    /**
     * 配列の中でのオブジェクトの位置を返す（DOM の参照は === で比べられる）
     * @param {Object[]} items - 探す配列
     * @param {Object} target - 探すオブジェクト
     * @returns {number} 位置。無ければ -1
     */
    function indexOfItem(items, target) {
        for (var i = 0; i < items.length; i++) {
            if (items[i] === target) return i;
        }
        return -1;
    }

    /**
     * 作業レイヤーを作って処理を実行し、終わったら作業レイヤーを消して選択とアクティブレイヤーを戻す
     * @param {Document} doc - 対象ドキュメント
     * @param {Function} work - (workLayer) を受け取る処理
     * @returns {void}
     */
    function withSymbolWorkLayer(doc, work) {
        var savedSelection = toItemArray(doc.selection);
        var savedActiveLayer = doc.activeLayer;
        var workLayer = doc.layers.add();
        try {
            work(workLayer);
        } finally {
            workLayer.remove();
            doc.activeLayer = savedActiveLayer;
            doc.selection = (savedSelection.length > 0) ? savedSelection : null;
        }
    }

    /**
     * シンボルの中身を作業レイヤーに展開し、1つのグループにまとめて返す
     * @param {Document} doc - 対象ドキュメント
     * @param {Symbol} symbol - 展開するシンボル
     * @param {Layer} workLayer - 作業レイヤー（空であること）
     * @returns {GroupItem} シンボルの中身をまとめたグループ
     */
    function expandSymbolToGroup(doc, symbol, workLayer) {
        doc.activeLayer = workLayer;
        var workInstance = doc.symbolItems.add(symbol);
        doc.selection = null;
        workInstance.selected = true;
        workInstance.breakLink();

        /* 解除の生成物はページアイテムかサブレイヤーのどちらかで出るので、作業レイヤーの中身を全部まとめる
           breakLink yields page items or a sublayer, so gather everything on the work layer */
        var contentGroup = workLayer.groupItems.add();
        var i;
        for (i = workLayer.pageItems.length - 1; i >= 0; i--) {
            if (workLayer.pageItems[i] !== contentGroup) workLayer.pageItems[i].move(contentGroup, ElementPlacement.PLACEATEND);
        }
        for (i = workLayer.layers.length - 1; i >= 0; i--) {
            var subLayer = workLayer.layers[i];
            while (subLayer.pageItems.length > 0) {
                subLayer.pageItems[0].move(contentGroup, ElementPlacement.PLACEATEND);
            }
            subLayer.remove();
        }
        return contentGroup;
    }

    /**
     * シンボル内のテキストの内容を集める（一致数の表示用。ドキュメントは変えない）
     * @param {Document} doc - 対象ドキュメント
     * @param {Symbol[]} targetSymbols - 対象のシンボル
     * @returns {string[]} テキストの内容
     */
    function getSymbolTextContents(doc, targetSymbols) {
        var textContents = [];
        if (targetSymbols.length === 0) return textContents;
        withSymbolWorkLayer(doc, function (workLayer) {
            for (var i = 0; i < targetSymbols.length; i++) {
                var contentGroup = expandSymbolToGroup(doc, targetSymbols[i], workLayer);
                var symbolFrames = [];
                collectItemsOfType(contentGroup, "TextFrame", symbolFrames);
                textContents = textContents.concat(getFrameContents(symbolFrames));
                contentGroup.remove();
            }
        });
        return textContents;
    }

    /**
     * シンボル内のテキストから一致箇所を削除し、書き換えたシンボルで定義を差し替える
     * @param {Document} doc - 対象ドキュメント
     * @param {Symbol[]} targetSymbols - 対象のシンボル
     * @param {Object[]} searchEntries - getSearchEntries() の結果
     * @param {number[]} removedCounts - 欄ごとの削除数（加算していく）
     * @returns {number} 書き換えたシンボルの数
     */
    function removeTextFromSymbols(doc, targetSymbols, searchEntries, removedCounts) {
        var updatedSymbolCount = 0;
        if (targetSymbols.length === 0) return updatedSymbolCount;
        withSymbolWorkLayer(doc, function (workLayer) {
            for (var i = 0; i < targetSymbols.length; i++) {
                var contentGroup = expandSymbolToGroup(doc, targetSymbols[i], workLayer);
                var symbolFrames = [];
                collectItemsOfType(contentGroup, "TextFrame", symbolFrames);
                /* 差し替えに失敗したら数えないよう、シンボルごとに数えてから足す / Count per symbol, and add only when the swap succeeds */
                var symbolRemovedCounts = [];
                for (var k = 0; k < searchEntries.length; k++) symbolRemovedCounts.push(0);
                var isChanged = false;
                for (var j = 0; j < symbolFrames.length; j++) {
                    if (removePatternsFromFrame(symbolFrames[j], searchEntries, symbolRemovedCounts)) isChanged = true;
                }
                if (isChanged && replaceSymbolDefinition(doc, targetSymbols[i], contentGroup)) {
                    updatedSymbolCount++;
                    for (k = 0; k < searchEntries.length; k++) removedCounts[k] += symbolRemovedCounts[k];
                }
                contentGroup.remove();
            }
        });
        return updatedSymbolCount;
    }

    /**
     * 新しいシンボルを作り、元のシンボルのインスタンスをすべて差し替えて、元のシンボルと入れ替える
     * （差し替えてもインスタンスの位置は変わらない。実測済み）
     * @param {Document} doc - 対象ドキュメント
     * @param {Symbol} oldSymbol - 元のシンボル
     * @param {GroupItem} contentGroup - 新しい定義にするアート
     * @returns {boolean} 差し替えられたら true
     */
    function replaceSymbolDefinition(doc, oldSymbol, contentGroup) {
        var newSymbol = doc.symbols.add(contentGroup);
        var swappedItems = [];
        /* 編集できないインスタンスがあれば、差し替えた分を戻して新しいシンボルを消す / If an instance cannot be edited, undo the swaps and drop the new symbol */
        try {
            for (var i = 0; i < doc.symbolItems.length; i++) {
                var symbolItem = doc.symbolItems[i];
                if (symbolItem.symbol !== oldSymbol) continue;
                var stateRecords = [];
                try {
                    unlockAndRevealAncestors(symbolItem, stateRecords);
                    symbolItem.symbol = newSymbol;
                    swappedItems.push(symbolItem);
                } finally {
                    restoreItemStates(stateRecords);
                }
            }
        } catch (e) {
            for (var j = 0; j < swappedItems.length; j++) swappedItems[j].symbol = oldSymbol;
            newSymbol.remove();
            return false;
        }

        var symbolName = oldSymbol.name;
        /* 別のシンボルの中から使われていると消せないので、そのときは元のシンボルを残す / An old symbol used inside another symbol cannot be removed, so keep it */
        try {
            oldSymbol.remove();
            newSymbol.name = symbolName;
        } catch (e) {}
        return true;
    }

    main();

})();
