#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

入力した文字列（5つまで、正規表現も可）を、選択中のオブジェクト・現在のアートボード・ドキュメント全体のテキストから削除、または別の文字列に置換します。
残った文字の書式は変わりません。シンボル内のテキストや、非表示・ロックされたテキストも対象にできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartTextFindReplace.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nec5dfffce709

### Overview

Removes up to five strings (regular expressions allowed) from text in the selection, the current artboard, or the entire document, or replaces them with other strings.
The formatting of the remaining text is kept. Text in symbols and hidden or locked text can be included.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartTextFindReplace.md

*/

(function () {

    // =========================================
    // 基本情報 / Basic info
    // =========================================
    var SCRIPT_NAME     = "SmartTextFindReplace";         /* スクリプト名 / script name */
    var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
    var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
    var SCRIPT_RELEASED = "2026-09-26";                   /* 最初のリリース日 / first release date */
    var SCRIPT_UPDATED  = "2026-09-26";                   /* 更新日 / last updated */

    var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartTextFindReplace.md"; /* README（日本語） */
    var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartTextFindReplace.md"; /* README (English) */
    var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nec5dfffce709"; /* 紹介記事 / article URL */

    // Released under the MIT license
    // http://opensource.org/licenses/mit-license.php

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var SEARCH_FIELD_COUNT = 5;             /* 削除・置換する文字列の入力欄の数 / number of search fields */
    var DEFAULT_USE_REGEX = false;          /* 正規表現（初期値）/ regular expression (initial value) */
    var DEFAULT_MATCH_CASE = true;          /* 大文字と小文字を区別（初期値）/ match case (initial value) */
    var DEFAULT_DELETE_EMPTY_FRAMES = true; /* 空になったフレームを削除（初期値）/ delete emptied frames (initial value) */
    var DEFAULT_INCLUDE_HIDDEN = false;     /* 非表示のレイヤーを検索（初期値）/ search hidden layers (initial value) */
    var DEFAULT_INCLUDE_LOCKED = false;     /* ロックされたレイヤーを検索（初期値）/ search locked layers (initial value) */
    var DEFAULT_INCLUDE_SYMBOLS = true;     /* シンボルも検索（初期値）/ search symbols too (initial value) */

    // =========================================
    // 改行の記号 / Break tokens
    // =========================================
    /* 1行の入力欄では改行を打てないので、記号で書いて処理の直前に文字へ置き換える
       Single-line fields cannot hold breaks, so they are written as tokens and converted before processing */
    var PARAGRAPH_BREAK_TOKEN = "\\n";      /* 改行（段落の区切り \r）/ paragraph break (\r) */
    var LINE_BREAK_TOKEN = "@#";            /* 強制改行（\x03）/ forced line break (\x03) */

    // =========================================
    // 前回の設定 / Saved settings
    // =========================================
    var SETTINGS_PREF_KEY = "SmartTextFindReplace/settings"; /* 環境設定に保存するキー / preference key */

    /**
     * 前回 OK したときの設定を読み込む
     * @returns {Object|null} 保存した設定。無いか壊れていれば null
     */
    function loadSettings() {
        var savedText = app.preferences.getStringPreference(SETTINGS_PREF_KEY);
        if (!savedText) return null;
        var savedSettings;
        /* 壊れた文字列は eval が例外を出す / eval throws on a corrupted string */
        try {
            savedSettings = eval(savedText);
        } catch (e) {
            return null;
        }
        /* v1.2.1 より前は「大文字と小文字を区別しない」（ignoreCase）で保存していたので、反転して読み替える
           Before v1.2.1 the setting was saved as ignoreCase, so read it inverted */
        if (savedSettings && typeof savedSettings.matchCase !== "boolean" && typeof savedSettings.ignoreCase === "boolean") {
            savedSettings.matchCase = !savedSettings.ignoreCase;
        }
        return savedSettings;
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
    var REPLACE_INPUT_CHARACTERS = 12;      /* 置換欄の幅（文字数）/ replace field width in characters */
    var MATCH_COUNT_WIDTH = 40;             /* 一致数の表示幅 / width of the match count */
    var SEARCH_OPTIONS_TOP_MARGIN = 5;      /* 正規表現などの上余白 / top margin above the search options */
    var INSERT_BUTTON_HEIGHT = 20;          /* 挿入ボタンの高さ / height of the insert buttons */
    var INSERT_BUTTON_FONT_SHRINK = 2;      /* 挿入ボタンの文字を小さくする量（pt）/ how much smaller the insert button font is (pt) */
    var INSERT_BUTTON_TOP_MARGIN = 2;       /* 挿入ボタンの上余白 / top margin of the insert buttons */
    var INSERT_BUTTON_LINE_SPACING = 4;     /* 挿入ボタンの行間 / spacing between rows of insert buttons */
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
            title: { ja: "指定した文字列を削除・置換", en: "Remove or Replace Text" }
        },
        panel: {
            findReplace: { ja: "削除・置換する文字列", en: "Text to Remove / Replace" },
            scope: { ja: "対象", en: "Scope" },
            options: { ja: "オプション", en: "Options" }
        },
        checkbox: {
            preview: { ja: "プレビュー", en: "Preview" },
            useRegex: { ja: "正規表現", en: "Regular expression" },
            matchCase: { ja: "大文字と小文字を区別", en: "Match case" },
            deleteEmptyFrames: { ja: "空になったテキストを削除", en: "Delete emptied text" },
            includeHidden: { ja: "非表示のレイヤーを検索", en: "Check hidden layers" },
            includeLocked: { ja: "ロックされたレイヤーを検索", en: "Check locked layers" },
            includeSymbols: { ja: "シンボルを検索", en: "Check symbols" }
        },
        radio: {
            selection: { ja: "選択中のオブジェクト", en: "Selected objects" },
            artboard: { ja: "現在のアートボード", en: "Current artboard" },
            wholeDocument: { ja: "ドキュメント全体", en: "Entire document" }
        },
        tooltip: {
            searchText: {
                ja: "一致する部分をすべて削除します（右の欄に入力すると置換）。空欄は無視し、上の欄から順に処理します。入力内容は次回に引き継がれます",
                en: "Removes every match (or replaces it when the field on the right is filled). Empty fields are ignored; fields are processed from top to bottom. Entries are kept for the next run"
            },
            replaceText: {
                ja: "一致した部分をこの文字列に置き換えます。空欄なら削除します。正規表現のときは $1・\\1 や $&・\\0 で一致した部分を参照できます",
                en: "Replaces each match with this text. Leave empty to remove. With regular expressions, $1 / \\1 and $& / \\0 refer to the match"
            },
            matchCount: {
                ja: "対象の範囲で見つかった数（ほかの欄による削除・置換は考慮しない）。「!」は正規表現の誤り",
                en: "Matches found in the scope (ignoring removals and replacements by other fields). \"!\" means an invalid regular expression"
            },
            paragraphBreak: {
                ja: "カーソルの位置に改行（\\n）を入れます（command／Ctrl＋Enter）",
                en: "Inserts a paragraph break (\\n) at the cursor (Command/Ctrl+Enter)"
            },
            lineBreak: {
                ja: "カーソルの位置に強制改行（@#）を入れます（Shift＋Enter）",
                en: "Inserts a forced line break (@#) at the cursor (Shift+Enter)"
            },
            wholeMatch: {
                ja: "カーソルの位置に \\0（一致した部分全体）を入れます（option／Alt＋0）。正規表現のときだけ使えます",
                en: "Inserts \\0 (the whole match) at the cursor (Option/Alt+0). Available with regular expressions only"
            },
            group1: {
                ja: "カーソルの位置に \\1（1つ目の ( ) に一致した部分）を入れます（option／Alt＋1）。正規表現のときだけ使えます",
                en: "Inserts \\1 (the text matched by the first group) at the cursor (Option/Alt+1). Available with regular expressions only"
            },
            group2: {
                ja: "カーソルの位置に \\2（2つ目の ( ) に一致した部分）を入れます（option／Alt＋2）。正規表現のときだけ使えます",
                en: "Inserts \\2 (the text matched by the second group) at the cursor (Option/Alt+2). Available with regular expressions only"
            },
            useRegex: {
                ja: "入力を JavaScript の正規表現として扱います。^ と $ は段落の先頭・末尾に一致します",
                en: "Treats the input as JavaScript regular expressions. ^ and $ match the start and end of each paragraph"
            },
            selection: { ja: "グループ内のテキストも対象になります", en: "Includes text inside groups" },
            artboard: { ja: "アートボードに一部でも重なるテキストが対象になります", en: "Includes text that partly overlaps the artboard" },
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
            selectionUnavailable: { ja: "オブジェクトが選択されていないため選べません", en: "Unavailable because nothing is selected" },
            reset: {
                ja: "入力欄を空にし、対象とオプションを初期値に戻します（プレビューの ON／OFF はそのまま）",
                en: "Clears the fields and restores the scope and options to their defaults (the preview setting is kept)"
            },
            preview: {
                ja: "ダイアログを閉じずに結果を表示し、入力や対象の変更に合わせて更新します。シンボル内・非表示・ロック中・スレッドテキスト（連結）のテキストは表示しません",
                en: "Shows the result without closing the dialog, updating as the input or scope changes. Text in symbols, hidden, locked or threaded text is not previewed"
            }
        },
        button: {
            paragraphBreak: { ja: "改行", en: "Paragraph Break" },
            lineBreak: { ja: "強制改行", en: "Forced Line Break" },
            wholeMatch: { ja: "検索結果すべて", en: "Whole Match" },
            group1: { ja: "検索結果1", en: "Group 1" },
            group2: { ja: "検索結果2", en: "Group 2" },
            reset: { ja: "リセット", en: "Reset" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            processedCounts: { ja: "削除・置換した数", en: "Removed / replaced matches" },
            removedCountLine: { ja: "「{text}」：{count}", en: "\"{text}\": {count}" },
            replacedCountLine: { ja: "「{text}」→「{replaceText}」：{count}", en: "\"{text}\" → \"{replaceText}\": {count}" },
            changedFrames: { ja: "{count}個のテキストオブジェクトを変更しました。", en: "Changed {count} text object(s)." },
            deletedFrames: { ja: "空になった{count}個のテキストフレームを削除しました。", en: "Deleted {count} text frame(s) left empty." },
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

        var findReplaceOptions = showFindReplaceDialog(doc, selectedItems);
        if (!findReplaceOptions) return;

        var targetFrames = collectTargetFrames(doc, findReplaceOptions.scope, findReplaceOptions.layerOptions, selectedItems);
        /* 空になったフレームを削除すると selectedItems に無効な参照が残るので、シンボルは先に集める
           Collect symbols first, as deleting emptied frames leaves invalid references in selectedItems */
        var targetSymbols = findReplaceOptions.includeSymbols ? collectTargetSymbols(doc, findReplaceOptions.scope, findReplaceOptions.layerOptions, selectedItems) : [];
        var replaceResult = replaceTextInFrames(targetFrames, findReplaceOptions.searchEntries, findReplaceOptions.deleteEmptyFrames);
        replaceResult.updatedSymbolCount = replaceTextInSymbols(doc, targetSymbols, findReplaceOptions.searchEntries, replaceResult.processedCounts);
        alert(buildResultMessage(findReplaceOptions.searchEntries, replaceResult));
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
     * 結果メッセージを組み立てる（欄ごとの削除・置換数・変更したオブジェクト数・削除したフレーム数）
     * @param {Object[]} searchEntries - getSearchEntries() の結果
     * @param {Object} replaceResult - replaceTextInFrames() の結果
     * @returns {string} メッセージ
     */
    function buildResultMessage(searchEntries, replaceResult) {
        var messageLines = [labelText(LABELS.alert.processedCounts)];
        for (var i = 0; i < searchEntries.length; i++) {
            var countLine = (searchEntries[i].replaceText === "") ? LABELS.alert.removedCountLine : LABELS.alert.replacedCountLine;
            messageLines.push(getLabel(countLine)
                .replace("{text}", searchEntries[i].text)
                .replace("{replaceText}", searchEntries[i].replaceText)
                .replace("{count}", replaceResult.processedCounts[i]));
        }
        messageLines.push("");
        messageLines.push(getLabel(LABELS.alert.changedFrames).replace("{count}", replaceResult.changedCount));
        if (replaceResult.deletedFrameCount > 0) {
            messageLines.push(getLabel(LABELS.alert.deletedFrames).replace("{count}", replaceResult.deletedFrameCount));
        }
        if (replaceResult.updatedSymbolCount > 0) {
            messageLines.push(getLabel(LABELS.alert.updatedSymbols).replace("{count}", replaceResult.updatedSymbolCount));
        }
        return messageLines.join("\n");
    }

    /**
     * ダイアログを表示し、削除・置換するパターンと対象範囲を返す
     * @param {Document} doc - 一致数を数えるドキュメント
     * @param {PageItem[]} selectedItems - 実行時に選択していたアイテム
     * @returns {Object|null} { searchEntries: Object[], scope: string, deleteEmptyFrames: boolean, includeSymbols: boolean, layerOptions: Object }。キャンセル時は null
     */
    function showFindReplaceDialog(doc, selectedItems) {
        /* 前回の設定があれば初期値にする / Start from the saved settings when present */
        var savedSettings = loadSettings() || {};
        var hasSelection = selectedItems.length > 0;
        /* チェックボックスと対象の現在値。show() 前は checkbox.value を読み戻せないので、ここに持つ
           Current values of the checkboxes and scope, kept here as checkbox.value cannot be read back before show() */
        var dialogState = {
            values: {},
            settingControls: [],
            scope: getInitialScope(savedSettings.scope, hasSelection),
            onSettingChange: null
        };

        var dlg = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];

        var findReplaceControls = buildFindReplacePanel(dlg, savedSettings, dialogState);

        /* 対象とオプションを2カラムに並べる / Place the scope and options panels in two columns */
        var scopeOptionsGroup = dlg.add("group");
        scopeOptionsGroup.orientation = "row";
        scopeOptionsGroup.alignChildren = ["fill", "fill"];
        var scopeRadios = buildScopePanel(scopeOptionsGroup, dialogState, hasSelection);
        buildOptionsPanel(scopeOptionsGroup, savedSettings, dialogState);

        var buttonControls = buildButtonRow(dlg);

        var searchInputs = findReplaceControls.searchInputs;
        var replaceInputs = findReplaceControls.replaceInputs;
        /* 対象範囲ごとのテキスト内容（範囲やレイヤーの扱いを切り替えたときだけ集め直す）/ Contents cached per scope and layer options */
        var contentsCache = {};
        /* 表示中のプレビュー（createPreview() の記録）/ Records of the preview on screen */
        var previewRecords = [];
        /* 元を隠すと選択が外れるので、一度でもプレビューしたら閉じたあとに選択を戻す
           Hiding the originals deselects them, so restore the selection after closing once a preview was shown */
        var hasShownPreview = false;
        var isPreviewOn = false;
        /* 入力が有効か（有効なパターンがあり、誤った正規表現が無い）/ Whether the input is valid */
        var isInputValid = false;

        /* 非表示・ロックの扱い / How hidden and locked layers are handled */
        function getLayerOptions() {
            return { includeHidden: dialogState.values.includeHidden, includeLocked: dialogState.values.includeLocked };
        }

        /* 入力欄から検索エントリーを作る / Build the search entries from the fields */
        function readSearchEntries() {
            return getSearchEntries(readInputTexts(searchInputs), readInputTexts(replaceInputs), dialogState.values.useRegex, dialogState.values.matchCase);
        }

        /* プレビューを消して元に戻す / Remove the preview and restore the originals */
        function clearPreview() {
            if (previewRecords.length === 0) return;
            removePreview(previewRecords);
            previewRecords = [];
        }

        /* プレビューを作り直す（OFF か入力が無効なら消すだけ）/ Rebuild the preview; only clear it when off or the input is invalid */
        function updatePreview() {
            clearPreview();
            if (!isPreviewOn || !isInputValid) return;
            previewRecords = createPreview(collectTargetFrames(doc, dialogState.scope, getLayerOptions(), selectedItems), readSearchEntries());
            if (previewRecords.length > 0) hasShownPreview = true;
        }

        /* 対象範囲のテキスト内容（シンボル内を含む）/ Text contents in the scope, including symbols */
        function getScopeContents() {
            var layerOptions = getLayerOptions();
            var includeSymbols = dialogState.values.includeSymbols;
            var cacheKey = [dialogState.scope, layerOptions.includeHidden, layerOptions.includeLocked, includeSymbols].join(":");
            if (!contentsCache[cacheKey]) {
                var textContents = getFrameContents(collectTargetFrames(doc, dialogState.scope, layerOptions, selectedItems));
                if (includeSymbols) {
                    textContents = textContents.concat(getSymbolTextContents(doc, collectTargetSymbols(doc, dialogState.scope, layerOptions, selectedItems)));
                }
                contentsCache[cacheKey] = textContents;
            }
            return contentsCache[cacheKey];
        }

        /* 一致数と OK の可否、プレビューを更新（有効なパターンが無いか、誤った正規表現がある間は OK を押せない）
           Update match counts, OK and the preview; disable OK when no pattern is given or a regular expression is invalid */
        function refreshDialogState() {
            /* プレビューの複製が数に入らないよう、先に消す / Clear the preview first so its duplicates are not counted */
            clearPreview();
            var scopeContents = getScopeContents();
            var hasPattern = false;
            var hasInvalidPattern = false;
            for (var i = 0; i < searchInputs.length; i++) {
                var searchPattern = createSearchPattern(searchInputs[i].text, dialogState.values.useRegex, dialogState.values.matchCase);
                var matchCountText = "";
                if (searchPattern === false) {
                    matchCountText = "!";
                    hasInvalidPattern = true;
                } else if (searchPattern) {
                    matchCountText = String(countMatches(scopeContents, searchPattern));
                    hasPattern = true;
                }
                findReplaceControls.matchCountLabels[i].text = matchCountText;
            }
            isInputValid = hasPattern && !hasInvalidPattern;
            buttonControls.btnOK.enabled = isInputValid;
            findReplaceControls.referenceButtonGroup.enabled = dialogState.values.useRegex;
            updatePreview();
        }

        /* 入力欄を空にし、対象とオプションを初期値に戻す（プレビューの ON/OFF はそのまま）
           Clear the fields and restore the scope and options to their defaults, keeping the preview setting */
        function resetDialog() {
            for (var i = 0; i < searchInputs.length; i++) {
                searchInputs[i].text = "";
                replaceInputs[i].text = "";
            }
            for (var j = 0; j < dialogState.settingControls.length; j++) {
                var settingControl = dialogState.settingControls[j];
                settingControl.checkbox.value = settingControl.defaultValue;
                dialogState.values[settingControl.settingKey] = settingControl.defaultValue;
            }
            dialogState.scope = getInitialScope(null, hasSelection);
            scopeRadios[dialogState.scope].value = true;
            findReplaceControls.focusFirstInput();
            refreshDialogState();
        }

        for (var i = 0; i < searchInputs.length; i++) {
            searchInputs[i].onChanging = refreshDialogState;
            replaceInputs[i].onChanging = updatePreview;
        }
        dialogState.onSettingChange = refreshDialogState;
        buttonControls.btnReset.onClick = resetDialog;
        buttonControls.previewCheckbox.onClick = function () {
            isPreviewOn = this.value;
            updatePreview();
        };

        refreshDialogState();

        var dialogResult = dlg.show();
        /* OK でもキャンセルでも、本処理の前にプレビューを消す / Remove the preview before running, whether OK or cancel */
        clearPreview();
        if (hasShownPreview && hasSelection) {
            /* 戻せない選択（ロック・非表示になったものなど）は例外になるので無視する / Ignore a selection that can no longer be restored */
            try {
                doc.selection = selectedItems;
            } catch (e) {}
        }
        if (dialogResult !== 1) return null;

        var searchTexts = readInputTexts(searchInputs);
        var replaceTexts = readInputTexts(replaceInputs);
        var settingsToSave = { searchTexts: searchTexts, replaceTexts: replaceTexts, scope: dialogState.scope };
        for (var k = 0; k < dialogState.settingControls.length; k++) {
            var settingKey = dialogState.settingControls[k].settingKey;
            settingsToSave[settingKey] = dialogState.values[settingKey];
        }
        saveSettings(settingsToSave);

        return {
            searchEntries: getSearchEntries(searchTexts, replaceTexts, dialogState.values.useRegex, dialogState.values.matchCase),
            scope: dialogState.scope,
            deleteEmptyFrames: dialogState.values.deleteEmptyFrames,
            includeSymbols: dialogState.values.includeSymbols,
            layerOptions: getLayerOptions()
        };
    }

    /**
     * 最初に選ぶ対象範囲を決める（何も選択していなければ「選択中のオブジェクト」は選べない）
     * @param {string|null} savedScope - 前回の対象範囲
     * @param {boolean} hasSelection - 選択があるか
     * @returns {string} "selection" / "artboard" / "document"
     */
    function getInitialScope(savedScope, hasSelection) {
        if (savedScope === "artboard") return "artboard";
        if (savedScope === "document" || !hasSelection) return "document";
        return "selection";
    }

    /**
     * 入力欄の文字列を配列にする
     * @param {EditText[]} inputs - 入力欄
     * @returns {string[]} 各欄の文字列
     */
    function readInputTexts(inputs) {
        var inputTexts = [];
        for (var i = 0; i < inputs.length; i++) {
            inputTexts.push(inputs[i].text);
        }
        return inputTexts;
    }

    /**
     * 設定を保存するチェックボックスを追加する（前回の値か初期値で始め、リセットと保存のために記録する）
     * @param {Object} parentGroup - 追加先のパネルかグループ
     * @param {string} settingKey - LABELS.checkbox・LABELS.tooltip・保存する設定に共通のキー
     * @param {boolean} defaultValue - 初期値
     * @param {Object} savedSettings - loadSettings() の結果
     * @param {Object} dialogState - showFindReplaceDialog() の状態
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addSettingCheckbox(parentGroup, settingKey, defaultValue, savedSettings, dialogState) {
        var settingCheckbox = parentGroup.add("checkbox", undefined, getLabel(LABELS.checkbox[settingKey]));
        var initialValue = (typeof savedSettings[settingKey] === "boolean") ? savedSettings[settingKey] : defaultValue;
        settingCheckbox.value = initialValue;
        if (LABELS.tooltip[settingKey]) settingCheckbox.helpTip = getLabel(LABELS.tooltip[settingKey]);
        dialogState.values[settingKey] = initialValue;
        dialogState.settingControls.push({ settingKey: settingKey, checkbox: settingCheckbox, defaultValue: defaultValue });
        settingCheckbox.onClick = function () {
            dialogState.values[settingKey] = this.value;
            if (dialogState.onSettingChange) dialogState.onSettingChange();
        };
        return settingCheckbox;
    }

    /**
     * 「削除・置換する文字列」パネルを作る（入力欄・挿入ボタン・正規表現などのチェックボックス）
     * @param {Window} dlg - ダイアログ
     * @param {Object} savedSettings - loadSettings() の結果
     * @param {Object} dialogState - showFindReplaceDialog() の状態
     * @returns {Object} { searchInputs: EditText[], replaceInputs: EditText[], matchCountLabels: StaticText[], referenceButtonGroup: Group, focusFirstInput: Function }
     */
    function buildFindReplacePanel(dlg, savedSettings, dialogState) {
        var savedSearchTexts = savedSettings.searchTexts || [];
        var savedReplaceTexts = savedSettings.replaceTexts || [];
        var findReplacePanel = dlg.add("panel", undefined, getLabel(LABELS.panel.findReplace));
        setupPanel(findReplacePanel);

        var searchInputs = [];
        var replaceInputs = [];
        var matchCountLabels = [];
        /* 最後にカーソルがあった入力欄（挿入ボタンの挿入先）/ Field that last had the cursor, where the insert buttons insert */
        var lastActiveInput = null;
        /* ボタンを押すと入力欄のカーソルが失われるので、押し下げた時点の位置を控える / Clicking a button loses the caret, so keep its position from the mouse down */
        var savedCaret = null;

        /* 入力欄に、挿入先の記録とショートカットを付ける / Track the cursor field and add the shortcuts */
        function setupShortcutInput(input) {
            input.onActivate = function () {
                lastActiveInput = this;
            };
            input.addEventListener("keydown", function (event) {
                var insertedToken = getShortcutToken(event.keyName, ScriptUI.environment.keyboardState, dialogState.values.useRegex);
                if (!insertedToken) return;
                /* Enter で OK が押されたり、option＋数字で記号が入ったりしないよう止める
                   Keep Enter from pressing OK and option+digit from typing a symbol */
                event.preventDefault();
                this.textselection = insertedToken;
                if (this.onChanging) this.onChanging();
            });
        }

        for (var i = 0; i < SEARCH_FIELD_COUNT; i++) {
            var fieldRowGroup = findReplacePanel.add("group");
            fieldRowGroup.orientation = "row";
            fieldRowGroup.alignChildren = ["left", "center"];
            var searchInput = fieldRowGroup.add("edittext", undefined, savedSearchTexts[i] || "");
            searchInput.characters = INPUT_CHARACTERS;
            searchInput.helpTip = getLabel(LABELS.tooltip.searchText);
            setupShortcutInput(searchInput);
            fieldRowGroup.add("statictext", undefined, "→");
            var replaceInput = fieldRowGroup.add("edittext", undefined, savedReplaceTexts[i] || "");
            replaceInput.characters = REPLACE_INPUT_CHARACTERS;
            replaceInput.helpTip = getLabel(LABELS.tooltip.replaceText);
            setupShortcutInput(replaceInput);
            var matchCountLabel = fieldRowGroup.add("statictext", undefined, "");
            matchCountLabel.preferredSize.width = MATCH_COUNT_WIDTH;
            matchCountLabel.justify = "right";
            matchCountLabel.helpTip = getLabel(LABELS.tooltip.matchCount);
            searchInputs.push(searchInput);
            replaceInputs.push(replaceInput);
            matchCountLabels.push(matchCountLabel);
        }

        /* 先頭の入力欄にカーソルを置く / Put the cursor in the first field */
        function focusFirstInput() {
            lastActiveInput = searchInputs[0];
            searchInputs[0].active = true;
        }
        focusFirstInput();

        /* 挿入ボタンの行（左・スペーサー・右の3カラム）/ Row of insert buttons: left, spacer and right columns */
        var insertButtonRowGroup = findReplacePanel.add("group");
        insertButtonRowGroup.orientation = "row";
        insertButtonRowGroup.alignment = ["fill", "top"];
        insertButtonRowGroup.alignChildren = ["left", "top"];
        insertButtonRowGroup.margins = [0, INSERT_BUTTON_TOP_MARGIN, 0, 0];

        /* 押すとカーソルの位置に記号を入れるボタンを作る / Create a button that inserts a token at the cursor */
        function addInsertButton(parentGroup, labelKey, insertedToken) {
            var insertButton = addSmallButton(parentGroup, labelKey);
            /* クリックが確定する前（押し下げた時点）にカーソル位置を読む / Read the caret on mouse down, before the click completes */
            insertButton.addEventListener("mousedown", function () {
                savedCaret = captureCaret(lastActiveInput);
            });
            insertButton.onClick = function () {
                insertTokenAtCaret(lastActiveInput, insertedToken, savedCaret);
                if (lastActiveInput.onChanging) lastActiveInput.onChanging();
            };
        }

        /* 左：改行 / Left: breaks */
        var breakButtonGroup = addButtonRowGroup(insertButtonRowGroup);
        addInsertButton(breakButtonGroup, "paragraphBreak", PARAGRAPH_BREAK_TOKEN);
        addInsertButton(breakButtonGroup, "lineBreak", LINE_BREAK_TOKEN);

        /* 中央：スペーサー（伸縮）/ Center: spacer (stretchable) */
        var insertButtonSpacer = insertButtonRowGroup.add("group");
        insertButtonSpacer.alignment = ["fill", "fill"];
        insertButtonSpacer.minimumSize.width = 0;

        /* 右：検索結果の参照（正規表現のときだけ使える）。1行目に「検索結果すべて」、2行目に「検索結果1」「検索結果2」
           Right: match references, available with regular expressions only. Whole match on the first line, groups 1 and 2 on the second */
        var referenceButtonGroup = insertButtonRowGroup.add("group");
        referenceButtonGroup.orientation = "column";
        referenceButtonGroup.alignment = ["right", "top"];
        referenceButtonGroup.alignChildren = ["left", "top"];
        referenceButtonGroup.spacing = INSERT_BUTTON_LINE_SPACING;
        var wholeMatchButtonGroup = addButtonRowGroup(referenceButtonGroup);
        addInsertButton(wholeMatchButtonGroup, "wholeMatch", "\\0");
        var groupReferenceButtonGroup = addButtonRowGroup(referenceButtonGroup);
        addInsertButton(groupReferenceButtonGroup, "group1", "\\1");
        addInsertButton(groupReferenceButtonGroup, "group2", "\\2");

        /* 正規表現・大文字と小文字 / Regular expression and case */
        var searchOptionsGroup = findReplacePanel.add("group");
        searchOptionsGroup.orientation = "column";
        searchOptionsGroup.alignChildren = ["left", "top"];
        searchOptionsGroup.margins = [0, SEARCH_OPTIONS_TOP_MARGIN, 0, 0];
        searchOptionsGroup.spacing = PANEL_SPACING;
        addSettingCheckbox(searchOptionsGroup, "useRegex", DEFAULT_USE_REGEX, savedSettings, dialogState);
        addSettingCheckbox(searchOptionsGroup, "matchCase", DEFAULT_MATCH_CASE, savedSettings, dialogState);

        return {
            searchInputs: searchInputs,
            replaceInputs: replaceInputs,
            matchCountLabels: matchCountLabels,
            referenceButtonGroup: referenceButtonGroup,
            focusFirstInput: focusFirstInput
        };
    }

    /**
     * ボタンを横に並べるグループを作る
     * @param {Object} parentGroup - 追加先のグループ
     * @returns {Group} 作ったグループ
     */
    function addButtonRowGroup(parentGroup) {
        var buttonRowGroup = parentGroup.add("group");
        buttonRowGroup.orientation = "row";
        buttonRowGroup.alignChildren = ["left", "center"];
        return buttonRowGroup;
    }

    /**
     * ほかのボタンよりひとまわり小さいボタンを作る
     * @param {Object} parentGroup - 追加先のグループ
     * @param {string} labelKey - LABELS.button と LABELS.tooltip のキー
     * @returns {Button} 作ったボタン
     */
    function addSmallButton(parentGroup, labelKey) {
        var smallButton = parentGroup.add("button", undefined, getLabel(LABELS.button[labelKey]));
        smallButton.helpTip = getLabel(LABELS.tooltip[labelKey]);
        var buttonFont = smallButton.graphics.font;
        smallButton.graphics.font = ScriptUI.newFont(buttonFont.name, buttonFont.style, buttonFont.size - INSERT_BUTTON_FONT_SHRINK);
        smallButton.preferredSize.height = INSERT_BUTTON_HEIGHT;
        return smallButton;
    }

    /**
     * 「対象」パネルを作る
     * @param {Object} parentGroup - 追加先のグループ
     * @param {Object} dialogState - showFindReplaceDialog() の状態（scope を読み書きする）
     * @param {boolean} hasSelection - 選択があるか
     * @returns {Object} 対象範囲の名前をキーにしたラジオボタン { selection, artboard, document }
     */
    function buildScopePanel(parentGroup, dialogState, hasSelection) {
        var scopePanel = parentGroup.add("panel", undefined, getLabel(LABELS.panel.scope));
        setupPanel(scopePanel);
        var scopeRadios = {
            selection: scopePanel.add("radiobutton", undefined, getLabel(LABELS.radio.selection)),
            artboard: scopePanel.add("radiobutton", undefined, getLabel(LABELS.radio.artboard)),
            document: scopePanel.add("radiobutton", undefined, getLabel(LABELS.radio.wholeDocument))
        };
        scopeRadios.selection.helpTip = getLabel(LABELS.tooltip.selection);
        scopeRadios.artboard.helpTip = getLabel(LABELS.tooltip.artboard);
        if (!hasSelection) {
            scopeRadios.selection.enabled = false;
            scopeRadios.selection.helpTip = getLabel(LABELS.tooltip.selectionUnavailable);
        }
        scopeRadios[dialogState.scope].value = true;

        /* クリックした範囲を控えて更新する / Keep the clicked scope and refresh */
        function selectScopeOnClick(scopeName) {
            scopeRadios[scopeName].onClick = function () {
                dialogState.scope = scopeName;
                if (dialogState.onSettingChange) dialogState.onSettingChange();
            };
        }
        selectScopeOnClick("selection");
        selectScopeOnClick("artboard");
        selectScopeOnClick("document");
        return scopeRadios;
    }

    /**
     * 「オプション」パネルを作る
     * @param {Object} parentGroup - 追加先のグループ
     * @param {Object} savedSettings - loadSettings() の結果
     * @param {Object} dialogState - showFindReplaceDialog() の状態
     * @returns {void}
     */
    function buildOptionsPanel(parentGroup, savedSettings, dialogState) {
        var optionsPanel = parentGroup.add("panel", undefined, getLabel(LABELS.panel.options));
        setupPanel(optionsPanel);
        addSettingCheckbox(optionsPanel, "deleteEmptyFrames", DEFAULT_DELETE_EMPTY_FRAMES, savedSettings, dialogState);
        addSettingCheckbox(optionsPanel, "includeHidden", DEFAULT_INCLUDE_HIDDEN, savedSettings, dialogState);
        addSettingCheckbox(optionsPanel, "includeLocked", DEFAULT_INCLUDE_LOCKED, savedSettings, dialogState);
        addSettingCheckbox(optionsPanel, "includeSymbols", DEFAULT_INCLUDE_SYMBOLS, savedSettings, dialogState);
    }

    /**
     * ボタンエリアを作る（左にプレビューとリセット、右にキャンセルと OK）
     * @param {Window} dlg - ダイアログ
     * @returns {Object} { previewCheckbox: Checkbox, btnReset: Button, btnOK: Button }
     */
    function buildButtonRow(dlg) {
        var btnRowGroup = dlg.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var previewCheckbox = btnLeftGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.preview));
        previewCheckbox.helpTip = getLabel(LABELS.tooltip.preview);
        var btnReset = btnLeftGroup.add("button", undefined, getLabel(LABELS.button.reset));
        btnReset.helpTip = getLabel(LABELS.tooltip.reset);

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        return { previewCheckbox: previewCheckbox, btnReset: btnReset, btnOK: btnOK };
    }

    /**
     * ショートカットに対応する記号を返す
     * command（Ctrl）＋Enter：改行、Shift＋Enter：強制改行、option（Alt）＋0〜2：検索結果（正規表現のときだけ）
     * @param {string} keyName - 押されたキー
     * @param {Object} keyboardState - ScriptUI.environment.keyboardState
     * @param {boolean} useRegex - 正規表現が ON か
     * @returns {string|null} 入れる記号。該当しなければ null
     */
    function getShortcutToken(keyName, keyboardState, useRegex) {
        if (keyName === "Enter") {
            if (keyboardState.metaKey || keyboardState.ctrlKey) return PARAGRAPH_BREAK_TOKEN;
            if (keyboardState.shiftKey) return LINE_BREAK_TOKEN;
            return null;
        }
        if (keyboardState.altKey && useRegex && (keyName === "0" || keyName === "1" || keyName === "2")) {
            return "\\" + keyName;
        }
        return null;
    }

    /**
     * 控えたカーソルの位置に記号を入れる（ボタン用）。位置が取れていなければ末尾に足す
     * @param {EditText} input - 入れる先の入力欄
     * @param {string} insertedToken - 入れる記号
     * @param {Object|null} caret - captureCaret() の結果
     * @returns {void}
     */
    function insertTokenAtCaret(input, insertedToken, caret) {
        var currentText = input.text;
        if (caret && caret.input === input && caret.start + caret.length <= currentText.length) {
            input.text = currentText.substring(0, caret.start) + insertedToken + currentText.substring(caret.start + caret.length);
        } else {
            input.text = currentText + insertedToken;
        }
    }

    /**
     * 入力欄のカーソル位置を読む。目印の文字を選択範囲に差し込んで位置を測り、元の文字列に戻す
     * （ScriptUI にはカーソル位置を返すプロパティが無いため）
     * @param {EditText} input - 対象の入力欄（カーソルがあるうちに呼ぶ）
     * @returns {Object|null} { input: EditText, start: number, length: number }。読めなければ null
     */
    function captureCaret(input) {
        var CARET_MARKER = "\u0001";
        var originalText = input.text;
        var selectedLength = input.textselection.length;
        input.textselection = CARET_MARKER;
        var markerIndex = input.text.indexOf(CARET_MARKER);
        input.text = originalText;
        /* カーソルを失った入力欄では目印が先頭に入るので、先頭は信用しない（末尾に足す側に倒す）
           A field that lost its caret puts the marker at the start, so treat the start as unknown */
        if (markerIndex < 0 || (markerIndex === 0 && originalText.length > 0)) return null;
        return { input: input, start: markerIndex, length: selectedLength };
    }

    /**
     * 対象のテキストフレームを複製して削除・置換し、元を一時的に隠す（プレビュー用）
     * 表示中でロックされていない独立フレームのうち、一致があるものだけを扱う
     * @param {TextFrame[]} targetFrames - 対象のテキストフレーム
     * @param {Object[]} searchEntries - getSearchEntries() の結果
     * @returns {Object[]} { originalFrame: TextFrame, previewFrame: TextFrame, isOriginalHidden: boolean } の配列
     */
    function createPreview(targetFrames, searchEntries) {
        var previewRecords = [];
        var visibleOnly = { includeHidden: false, includeLocked: false };
        var ignoredCounts = createZeroCounts(searchEntries.length);
        for (var i = 0; i < targetFrames.length; i++) {
            var textFrame = targetFrames[i];
            if (!isSearchableItem(textFrame, visibleOnly) || !isStandaloneFrame(textFrame)) continue;
            if (!hasAnyMatch(textFrame.contents, searchEntries)) continue;
            /* 編集できないテキストは例外になるのでスキップ / Skip text that throws because it cannot be edited */
            try {
                /* 複製は元を隠す前に作る（hidden を引き継ぐため）/ Duplicate before hiding, as duplicates inherit hidden */
                var previewRecord = { originalFrame: textFrame, previewFrame: textFrame.duplicate(textFrame, ElementPlacement.PLACEBEFORE), isOriginalHidden: false };
                previewRecords.push(previewRecord);
                textFrame.hidden = true;
                previewRecord.isOriginalHidden = true;
                replacePatternsInFrame(previewRecord.previewFrame, searchEntries, ignoredCounts);
            } catch (e) {}
        }
        app.redraw();
        return previewRecords;
    }

    /**
     * プレビューの複製を消し、隠した元のフレームを表示に戻す
     * @param {Object[]} previewRecords - createPreview() の結果
     * @returns {void}
     */
    function removePreview(previewRecords) {
        for (var i = previewRecords.length - 1; i >= 0; i--) {
            previewRecords[i].previewFrame.remove();
            if (previewRecords[i].isOriginalHidden) previewRecords[i].originalFrame.hidden = false;
        }
        app.redraw();
    }

    /**
     * 文字列にいずれかのパターンが一致するか判定する
     * @param {string} text - 検索する文字列
     * @param {Object[]} searchEntries - getSearchEntries() の結果
     * @returns {boolean} 1つでも一致すれば true
     */
    function hasAnyMatch(text, searchEntries) {
        for (var i = 0; i < searchEntries.length; i++) {
            if (findMatches(text, searchEntries[i].pattern).length > 0) return true;
        }
        return false;
    }

    /**
     * 入力から検索パターンを作る
     * @param {string} searchText - 入力された文字列
     * @param {boolean} useRegex - 正規表現として扱うか
     * @param {boolean} matchCase - 大文字と小文字を区別するか
     * @returns {RegExp|null|boolean} パターン。空欄なら null、正規表現が誤っていれば false
     */
    function createSearchPattern(searchText, useRegex, matchCase) {
        if (searchText === "") return null;
        var source = useRegex ? convertBreakTokensInRegex(searchText) : expandBreakTokens(searchText).replace(/[.*+?^${}()|[\]\\\/]/g, "\\$&");
        /* 誤った正規表現は new RegExp() が例外を出す / new RegExp() throws on an invalid pattern */
        try {
            return new RegExp(source, matchCase ? "gm" : "gmi");
        } catch (e) {
            return false;
        }
    }

    /**
     * 改行の記号を文字に置き換える（\n は改行 \r、@# は強制改行 \x03）
     * @param {string} text - 入力された文字列
     * @returns {string} 置き換えた文字列
     */
    function expandBreakTokens(text) {
        return text.split(PARAGRAPH_BREAK_TOKEN).join("\r").split(LINE_BREAK_TOKEN).join("\x03");
    }

    /**
     * 正規表現の中の改行の記号を、Illustrator の改行文字に一致する書き方にする（\\ はそのまま）
     * @param {string} source - 入力された正規表現
     * @returns {string} 置き換えた正規表現
     */
    function convertBreakTokensInRegex(source) {
        return source.replace(/\\\\|\\n|@#/g, function (token) {
            if (token === PARAGRAPH_BREAK_TOKEN) return "\\r";
            if (token === LINE_BREAK_TOKEN) return "\\x03";
            return token;
        });
    }

    /**
     * 入力から有効な検索パターンを上から順に集める
     * @param {string[]} searchTexts - 入力欄の文字列
     * @param {string[]} replaceTexts - 置換欄の文字列（空欄は削除）
     * @param {boolean} useRegex - 正規表現として扱うか
     * @param {boolean} matchCase - 大文字と小文字を区別するか
     * @returns {Object[]} { text: string, pattern: RegExp, replaceText: string, useRegex: boolean } の配列
     */
    function getSearchEntries(searchTexts, replaceTexts, useRegex, matchCase) {
        var searchEntries = [];
        for (var i = 0; i < searchTexts.length; i++) {
            var searchPattern = createSearchPattern(searchTexts[i], useRegex, matchCase);
            if (searchPattern) {
                /* 正規表現の置換欄は getReplacementText() で改行の記号を展開する / Regex replacements expand the break tokens in getReplacementText() */
                var replaceText = replaceTexts[i] || "";
                searchEntries.push({ text: searchTexts[i], pattern: searchPattern, replaceText: useRegex ? replaceText : expandBreakTokens(replaceText), useRegex: useRegex });
            }
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
     * テキストフレームの一致箇所をすべて削除・置換する（パターンは配列の順に処理）
     * @param {TextFrame[]} targetFrames - 対象のテキストフレーム
     * @param {Object[]} searchEntries - getSearchEntries() の結果
     * @param {boolean} deleteEmptyFrames - 削除で空になった独立フレームを削除するか
     * @returns {Object} { processedCounts: number[]（欄ごとの削除・置換数）, changedCount: number, deletedFrameCount: number }
     */
    function replaceTextInFrames(targetFrames, searchEntries, deleteEmptyFrames) {
        var replaceResult = { processedCounts: createZeroCounts(searchEntries.length), changedCount: 0, deletedFrameCount: 0 };
        for (var i = 0; i < targetFrames.length; i++) {
            var textFrame = targetFrames[i];
            var stateRecords = [];
            /* 編集できないテキスト（テンプレートレイヤーなど）は例外になるのでスキップ / Skip text that throws because it cannot be edited (template layers, etc.) */
            try {
                unlockAndRevealAncestors(textFrame, stateRecords);
                var isChanged = replacePatternsInFrame(textFrame, searchEntries, replaceResult.processedCounts);
                if (isChanged) replaceResult.changedCount++;
                if (isChanged && deleteEmptyFrames && textFrame.contents === "" && isStandaloneFrame(textFrame)) {
                    /* 消したフレーム自身は状態を戻さない / Do not restore the deleted frame itself */
                    stateRecords = excludeRecordsFor(stateRecords, textFrame);
                    textFrame.remove();
                    replaceResult.deletedFrameCount++;
                }
            } catch (e) {
            } finally {
                restoreItemStates(stateRecords);
            }
        }
        return replaceResult;
    }

    /**
     * 欄ごとの数を数えるための、0 を並べた配列を作る
     * @param {number} length - 欄の数
     * @returns {number[]} 0 の配列
     */
    function createZeroCounts(length) {
        var zeroCounts = [];
        for (var i = 0; i < length; i++) zeroCounts.push(0);
        return zeroCounts;
    }

    /**
     * 1つのテキストフレームで、パターンを配列の順にすべて削除・置換する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object[]} searchEntries - getSearchEntries() の結果
     * @param {number[]} processedCounts - 欄ごとの削除・置換数（加算していく）
     * @returns {boolean} 1か所でも削除・置換したら true
     */
    function replacePatternsInFrame(textFrame, searchEntries, processedCounts) {
        var isChanged = false;
        for (var i = 0; i < searchEntries.length; i++) {
            /* 前の削除・置換で内容が変わるので、毎回 contents から探し直す / Search again each time, as earlier edits change the contents */
            var matches = findMatches(textFrame.contents, searchEntries[i].pattern);
            if (matches.length === 0) continue;
            replaceMatchesKeepingFormat(textFrame, matches, searchEntries[i]);
            processedCounts[i] += matches.length;
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
     * @returns {Object[]} { start: number, length: number, captures: Array } の配列（昇順）
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
            matches.push({ start: match.index, length: match[0].length, captures: match });
        }
        return matches;
    }

    /**
     * 一致箇所を右から削除・置換する（contents を書き戻さないので文字ごとの書式が残る）
     * 置換は先頭の1文字だけ残して contents を差し替えるので、置換後の文字はその文字の書式になる
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object[]} matches - findMatches() の結果（昇順）
     * @param {Object} searchEntry - getSearchEntries() の要素
     * @returns {void}
     */
    function replaceMatchesKeepingFormat(textFrame, matches, searchEntry) {
        var frameCharacters = textFrame.characters;
        for (var i = matches.length - 1; i >= 0; i--) {
            var newText = getReplacementText(searchEntry, matches[i].captures);
            var keptCount = (newText === "") ? 0 : 1;
            for (var j = matches[i].start + matches[i].length - 1; j >= matches[i].start + keptCount; j--) {
                frameCharacters[j].remove();
            }
            if (keptCount > 0) frameCharacters[matches[i].start].contents = newText;
        }
    }

    /**
     * 一致箇所の置換後の文字列を返す（正規表現のときは $$ / $& / $1〜$99 と \\ / \0〜\9 / \n / @# を展開する）
     * @param {Object} searchEntry - getSearchEntries() の要素
     * @param {Array} captures - exec() の結果
     * @returns {string} 置換後の文字列。空なら削除
     */
    function getReplacementText(searchEntry, captures) {
        if (!searchEntry.useRegex) return searchEntry.replaceText;
        return searchEntry.replaceText.replace(/\$(\$|&|\d\d?)|\\([\\\dn])|@#/g, function (token, dollarName, backslashName) {
            if (token === LINE_BREAK_TOKEN) return "\x03";
            /* \\ は \ そのもの、\n は改行、\0 は一致全体、\1〜\9 はグループ / \\ is a backslash, \n a paragraph break, \0 the whole match, \1-\9 the groups */
            if (backslashName) {
                if (backslashName === "\\") return "\\";
                if (backslashName === "n") return "\r";
                var backslashIndex = parseInt(backslashName, 10);
                if (backslashIndex >= captures.length) return token;
                return captures[backslashIndex] || "";
            }
            if (dollarName === "$") return "$";
            if (dollarName === "&") return captures[0];
            var groupIndex = parseInt(dollarName, 10);
            /* $12 でグループが足りなければ $1 と「2」として扱う / Read $12 as $1 followed by "2" when there are fewer groups */
            if (dollarName.length === 2 && groupIndex >= captures.length) {
                groupIndex = parseInt(dollarName.charAt(0), 10);
                if (groupIndex === 0 || groupIndex >= captures.length) return token;
                return (captures[groupIndex] || "") + dollarName.charAt(1);
            }
            if (groupIndex === 0 || groupIndex >= captures.length) return token;
            return captures[groupIndex] || "";
        });
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
     * @param {Function} workCallback - (workLayer) を受け取る処理
     * @returns {void}
     */
    function withSymbolWorkLayer(doc, workCallback) {
        var savedSelection = toItemArray(doc.selection);
        var savedActiveLayer = doc.activeLayer;
        var workLayer = doc.layers.add();
        try {
            workCallback(workLayer);
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
     * シンボル内のテキストの一致箇所を削除・置換し、書き換えたシンボルで定義を差し替える
     * @param {Document} doc - 対象ドキュメント
     * @param {Symbol[]} targetSymbols - 対象のシンボル
     * @param {Object[]} searchEntries - getSearchEntries() の結果
     * @param {number[]} processedCounts - 欄ごとの削除・置換数（加算していく）
     * @returns {number} 書き換えたシンボルの数
     */
    function replaceTextInSymbols(doc, targetSymbols, searchEntries, processedCounts) {
        var updatedSymbolCount = 0;
        if (targetSymbols.length === 0) return updatedSymbolCount;
        withSymbolWorkLayer(doc, function (workLayer) {
            for (var i = 0; i < targetSymbols.length; i++) {
                var contentGroup = expandSymbolToGroup(doc, targetSymbols[i], workLayer);
                var symbolFrames = [];
                collectItemsOfType(contentGroup, "TextFrame", symbolFrames);
                /* 差し替えに失敗したら数えないよう、シンボルごとに数えてから足す / Count per symbol, and add only when the swap succeeds */
                var symbolProcessedCounts = createZeroCounts(searchEntries.length);
                var isChanged = false;
                for (var j = 0; j < symbolFrames.length; j++) {
                    if (replacePatternsInFrame(symbolFrames[j], searchEntries, symbolProcessedCounts)) isChanged = true;
                }
                if (isChanged && replaceSymbolDefinition(doc, targetSymbols[i], contentGroup)) {
                    updatedSymbolCount++;
                    for (var k = 0; k < searchEntries.length; k++) processedCounts[k] += symbolProcessedCounts[k];
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
