#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

入力した文字列（7つまで、正規表現も可）を、選択中のオブジェクト・現在のアートボード・ドキュメント全体のテキストから削除、または別の文字列に置換します。
残った文字の書式は変わりません。シンボル内のテキストや、非表示・ロックされたテキストも対象にできます。
「英文」タブでは英字の大文字・小文字を変換し、「整形」タブではタブ・スペース・記号・かな・数字・行頭の箇条書きや番号を整えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartTextFindReplace.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nec5dfffce709

### Overview

Removes up to seven strings (regular expressions allowed) from text in the selection, the current artboard, or the entire document, or replaces them with other strings.
The formatting of the remaining text is kept. Text in symbols and hidden or locked text can be included.
The English tab changes letter case; the Cleanup tab tidies tabs, spaces, symbols, kana, digits, and leading bullets or numbers.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartTextFindReplace.md

*/

(function () {

    // =========================================
    // 基本情報 / Basic info
    // =========================================
    var SCRIPT_NAME     = "SmartTextFindReplace";         /* スクリプト名 / script name */
    var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
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
    var SEARCH_FIELD_COUNT = 7;             /* 削除・置換する文字列の入力欄の数 / number of search fields */
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
     * @param {Object} settingsToSave - 保存する設定
     * @returns {void}
     */
    function saveSettings(settingsToSave) {
        app.preferences.setStringPreference(SETTINGS_PREF_KEY, settingsToSave.toSource());
    }

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS = [15, 20, 15, 10];   /* パネルの内側余白 / panel margins */
    var PANEL_SPACING = 6;                  /* パネル内の間隔 / panel spacing */
    var CLEANUP_PANEL_MARGINS = [10, 20, 10, 10]; /* 「整形」タブのパネルの内側余白（左右を詰める）/ margins of the Cleanup tab panels, narrower left and right */
    var TAB_MARGINS = [10, 15, 10, 10];     /* タブの内側余白 / tab margins */
    var INPUT_CHARACTERS = 18;              /* 入力欄の幅（文字数）/ input width in characters */
    var REPLACE_INPUT_CHARACTERS = 12;      /* 置換欄の幅（文字数）/ replace field width in characters */
    var MATCH_COUNT_WIDTH = 24;             /* 一致数の表示幅（2桁ほど）/ width of the match count, about two digits */
    var SEARCH_OPTIONS_TOP_MARGIN = 5;      /* 正規表現などの上余白 / top margin above the search options */
    var INSERT_BUTTON_HEIGHT = 20;          /* 挿入ボタンの高さ / height of the insert buttons */
    var INSERT_BUTTON_FONT_SHRINK = 2;      /* 挿入ボタンの文字を小さくする量（pt）/ how much smaller the insert button font is (pt) */
    var INSERT_BUTTON_TOP_MARGIN = 2;       /* 挿入ボタンの上余白 / top margin of the insert buttons */
    var INSERT_BUTTON_LINE_SPACING = 4;     /* 挿入ボタンの行間 / spacing between rows of insert buttons */
    var BUTTON_ROW_TOP_MARGIN = 6;          /* ボタン行の上余白 / top margin of the button row */
    var CONVERSION_BUTTON_SIZE = [130, 24]; /* 変換ボタンのサイズ / size of the conversion buttons */
    var CONVERSION_SAMPLE_SIZE = [90, 24];  /* 変換結果の見本の最小サイズ（パネル幅まで伸びる）/ minimum size of the conversion samples, stretched to the panel width */
    var CONVERSION_SAMPLE_MAX_CHARS = 60;   /* 変換結果の見本に表示する最大文字数 / maximum characters in a conversion sample */

    /**
     * パネルの共通設定
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number[]} [panelMargins] - 内側余白（省略時は PANEL_MARGINS）
     * @returns {void}
     */
    function setupPanel(targetPanel, panelMargins) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["left", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = panelMargins || PANEL_MARGINS;
        targetPanel.spacing = PANEL_SPACING;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "テキストの削除・置換・整形", en: "Remove, Replace & Clean Up Text" }
        },
        tab: {
            findReplace: { ja: "削除・置換", en: "Remove / Replace" },
            englishText: { ja: "英文", en: "English" },
            cleanup: { ja: "整形", en: "Cleanup" }
        },
        panel: {
            findReplace: { ja: "削除・置換する文字列", en: "Text to Remove / Replace" },
            scope: { ja: "対象", en: "Scope" },
            options: { ja: "オプション", en: "Options" },
            letterCase: { ja: "大文字/小文字", en: "Letter Case" },
            kanaDigitConversion: { ja: "かな・数字の変換", en: "Kana & Digits" },
            tabCharacter: { ja: "タブ", en: "Tabs" },
            removeSpace: { ja: "スペース削除", en: "Remove Spaces" },
            addSpace: { ja: "スペース追加", en: "Add Spaces" },
            symbolConversion: { ja: "スペースや記号の変換", en: "Convert Spaces & Symbols" },
            symbolBefore: { ja: "変換前", en: "Before" },
            symbolAfter: { ja: "変換後", en: "After" },
            removeList: { ja: "リストの除去", en: "Remove List Markers" }
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
            wholeDocument: { ja: "ドキュメント全体", en: "Entire document" },
            space: { ja: "スペース", en: "Space" },
            underscore: { ja: "アンダースコア", en: "Underscore" },
            hyphen: { ja: "ハイフン", en: "Hyphen" }
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
                ja: "今回の削除で空になったテキストだけを削除します。元から空のテキストと、スレッドテキスト（連結）は残します",
                en: "Deletes only text emptied by this run. Text that was already empty and threaded text are kept"
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
            caseWord: { ja: "単語ごとに先頭を大文字、残りを小文字にします", en: "Capitalizes the first letter of each word and lowercases the rest" },
            caseSentence: { ja: "文の先頭だけを大文字にし、残りを小文字にします", en: "Capitalizes only the first letter of each sentence" },
            caseTitle: {
                ja: "冠詞・前置詞・接続詞などを除き、単語の先頭を大文字にします",
                en: "Capitalizes each word except articles, prepositions and conjunctions"
            },
            toHiragana: { ja: "カタカナ（半角カナを含む）をひらがなにします", en: "Converts katakana, including halfwidth kana, to hiragana" },
            toKatakana: { ja: "ひらがなと半角カナを全角カタカナにします", en: "Converts hiragana and halfwidth kana to fullwidth katakana" },
            toHalfKana: { ja: "ひらがなとカタカナを半角カナにします", en: "Converts hiragana and katakana to halfwidth kana" },
            toHalfDigit: { ja: "全角数字と漢数字を半角数字にします", en: "Converts fullwidth digits and kanji numerals to halfwidth digits" },
            toFullDigit: { ja: "半角数字と漢数字を全角数字にします", en: "Converts halfwidth digits and kanji numerals to fullwidth digits" },
            removeTabs: { ja: "タブをすべて削除します", en: "Removes all tabs" },
            tabsToSpaces: { ja: "タブを半角スペースに置き換えます", en: "Replaces tabs with spaces" },
            trimSpaces: { ja: "各行の行頭・行末のスペースを削除します", en: "Removes leading and trailing spaces on each line" },
            cjkLatinSpaces: {
                ja: "和文と欧文の間のスペースを削除します（欧文単語間は保持）",
                en: "Removes spaces between CJK and Latin text (spaces between Latin words are kept)"
            },
            collapseSpaces: { ja: "連続したスペースを1つにまとめます", en: "Collapses consecutive spaces into one" },
            cleanupSpaces: {
                ja: "行頭行末・連続・和欧間のスペースをまとめて処理します",
                en: "Applies Leading/Trailing, Consecutive and Between CJK and Latin in one step"
            },
            spaceAfterPunct: { ja: "半角ピリオド・カンマの直後にスペースを挿入します", en: "Inserts a space right after a period or comma" },
            convertSymbol: {
                ja: "変換前の記号を変換後の記号に置き換えます（スペースは半角・全角の両方）",
                en: "Replaces the Before symbol with the After symbol (spaces include fullwidth spaces)"
            },
            bulletList: {
                ja: "行頭の箇条書き記号（・ ･ · • ◦ ● ○ ◎ □ ■ ▪ ◆ ◇ ✓ – - *）を削除します。Illustrator の箇条書き機能は［テキストに変換］してから取り除きます",
                en: "Removes leading bullet markers (・ ･ · • ◦ ● ○ ◎ □ ■ ▪ ◆ ◇ ✓ – - *). Illustrator bullet lists are converted to text first, then removed"
            },
            numberList: {
                ja: "行頭の番号（1. ① a. 一. など）を削除します。Illustrator の番号付きリストは［テキストに変換］してから取り除きます",
                en: "Removes leading numbers (1. ① a. etc.). Illustrator numbered lists are converted to text first, then removed"
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
            caseUpper: { ja: "すべて大文字に", en: "UPPERCASE" },
            caseLower: { ja: "すべて小文字に", en: "lowercase" },
            caseWord: { ja: "単語の先頭を大文字", en: "Capitalize Words" },
            caseSentence: { ja: "文頭のみ大文字", en: "Sentence case" },
            caseTitle: { ja: "英語タイトル形式", en: "Title Case" },
            toHiragana: { ja: "ひらがな", en: "Hiragana" },
            toKatakana: { ja: "カタカナ", en: "Katakana" },
            toHalfKana: { ja: "半角カナ", en: "Halfwidth Kana" },
            toHalfDigit: { ja: "半角数字", en: "Halfwidth Digits" },
            toFullDigit: { ja: "全角数字", en: "Fullwidth Digits" },
            removeTabs: { ja: "削除", en: "Remove" },
            tabsToSpaces: { ja: "スペースに", en: "To Spaces" },
            trimSpaces: { ja: "行頭行末", en: "Leading/Trailing" },
            cjkLatinSpaces: { ja: "和欧間", en: "Between CJK and Latin" },
            collapseSpaces: { ja: "連続", en: "Consecutive" },
            cleanupSpaces: { ja: "まとめて", en: "All at Once" },
            spaceAfterPunct: { ja: ".と,の後", en: "After . and ," },
            convertSymbol: { ja: "変換", en: "Convert" },
            bulletList: { ja: "箇条書き", en: "Bullets" },
            numberList: { ja: "番号リスト", en: "Numbers" },
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
            deletedFrames: { ja: "空になった{count}個のテキストオブジェクトを削除しました。", en: "Deleted {count} emptied text object(s)." },
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
     * @param {Object} currentSelection - doc.selection
     * @returns {PageItem[]} アイテムの配列。選択が無ければ空
     */
    function toItemArray(currentSelection) {
        var itemArray = [];
        if (!currentSelection || !currentSelection.length) return itemArray;
        for (var i = 0; i < currentSelection.length; i++) {
            itemArray.push(currentSelection[i]);
        }
        return itemArray;
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
        /* 文字ツールで選択した文字（見本に使う）。シンボルの展開で選択が変わる前に控える
           Characters selected with the Type tool, used for the samples; kept before expanding symbols changes the selection */
        var selectedTextRange = (doc.selection && doc.selection.typename === "TextRange") ? doc.selection : null;

        var dialogControls = buildMainDialog(savedSettings, dialogState, hasSelection, runConversion);
        var findReplaceControls = dialogControls.findReplaceControls;
        var conversionPanelBuilder = dialogControls.conversionPanelBuilder;
        var buttonControls = dialogControls.buttonControls;

        var searchInputs = findReplaceControls.searchInputs;
        var replaceInputs = findReplaceControls.replaceInputs;
        /* 対象範囲ごとのテキスト内容（範囲やレイヤーの扱いを切り替えたときだけ集め直す）/ Contents cached per scope and layer options */
        var contentsCache = {};
        /* 表示中のプレビュー（createPreview() の記録）/ Records of the preview on screen */
        var previewRecords = [];
        /* 元を隠すと選択が外れるので、一度でもプレビューしたら閉じたあとに選択を戻す
           Hiding the originals deselects them, so restore the selection after closing once a preview was shown */
        var hasShownPreview = false;
        /* 箇条書きの変換で選択を使ったか / Whether converting list styles changed the selection */
        var hasChangedSelection = false;
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
            /* 見本は選択した文字・テキストを優先し、無ければ対象範囲の最初のテキスト / Samples prefer the selection, then the first text in the scope */
            conversionPanelBuilder.updateSamples(getSelectedSampleText(selectedTextRange, selectedItems) || getFirstNonEmptyText(scopeContents));
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

        /* 対象範囲のテキストをすぐに変換する（ダイアログは閉じない）/ Convert the text in the scope right away, keeping the dialog open */
        function runConversion(convertText, conversionKey) {
            /* プレビューの複製ではなく元を変換する / Convert the originals, not the preview duplicates */
            clearPreview();
            var layerOptions = getLayerOptions();
            var targetFrames = collectTargetFrames(doc, dialogState.scope, layerOptions, selectedItems);
            var targetSymbols = dialogState.values.includeSymbols ? collectTargetSymbols(doc, dialogState.scope, layerOptions, selectedItems) : [];
            if (conversionKey === "bulletList" || conversionKey === "numberList") {
                /* 箇条書き機能の記号を本文にしてから取り除く（選択を使うので、閉じたあとに選択を戻す）
                   Turn list markers into text first; this uses the selection, so restore it after closing */
                hasChangedSelection = true;
                var convertListStyle = function (textFrame) {
                    convertListStyleToText(doc, textFrame);
                };
                convertTextInFrames(targetFrames, convertText, convertListStyle);
                convertTextInSymbols(doc, targetSymbols, convertText, convertListStyle);
            } else {
                convertTextInFrames(targetFrames, convertText, null);
                convertTextInSymbols(doc, targetSymbols, convertText, null);
            }
            /* 内容が変わったので一致数を数え直す / The contents changed, so count the matches again */
            contentsCache = {};
            app.redraw();
            refreshDialogState();
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
            dialogControls.scopeRadios[dialogState.scope].value = true;
            findReplaceControls.focusFirstInput();
            refreshDialogState();
        }

        for (var i = 0; i < searchInputs.length; i++) {
            searchInputs[i].onChanging = refreshDialogState;
            replaceInputs[i].onChanging = updatePreview;
        }
        dialogState.onSettingChange = refreshDialogState;
        findReplaceControls.btnReset.onClick = resetDialog;
        buttonControls.previewCheckbox.onClick = function () {
            isPreviewOn = this.value;
            updatePreview();
        };

        refreshDialogState();

        var dialogResult = dialogControls.mainDialog.show();
        /* OK でもキャンセルでも、本処理の前にプレビューを消す / Remove the preview before running, whether OK or cancel */
        clearPreview();
        if (hasShownPreview || hasChangedSelection) {
            /* 戻せない選択（ロック・非表示になったものなど）は例外になるので無視する / Ignore a selection that can no longer be restored */
            try {
                doc.selection = hasSelection ? selectedItems : null;
            } catch (e) {}
        }
        if (dialogResult !== 1) return null;

        var searchTexts = readInputTexts(searchInputs);
        var replaceTexts = readInputTexts(replaceInputs);
        saveDialogSettings(searchTexts, replaceTexts, dialogState);

        return {
            searchEntries: getSearchEntries(searchTexts, replaceTexts, dialogState.values.useRegex, dialogState.values.matchCase),
            scope: dialogState.scope,
            deleteEmptyFrames: dialogState.values.deleteEmptyFrames,
            includeSymbols: dialogState.values.includeSymbols,
            layerOptions: getLayerOptions()
        };
    }

    /**
     * ダイアログを組み立てる（タブ・対象とオプション・ボタン行）
     * @param {Object} savedSettings - loadSettings() の結果
     * @param {Object} dialogState - showFindReplaceDialog() の状態
     * @param {boolean} hasSelection - 選択があるか
     * @param {Function} onConvert - 英文・整形タブのボタンを押したときに (convertText, conversionKey) を受け取る処理
     * @returns {Object} { mainDialog: Window, findReplaceControls: Object, conversionPanelBuilder: Object, scopeRadios: Object, buttonControls: Object }
     */
    function buildMainDialog(savedSettings, dialogState, hasSelection, onConvert) {
        var mainDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        mainDialog.orientation = "column";
        mainDialog.alignChildren = ["fill", "top"];

        /* 「削除・置換」「英文」「整形」をタブで切り替える / Switch between the remove/replace, English and cleanup tabs */
        var modeTabbedPanel = mainDialog.add("tabbedpanel");
        modeTabbedPanel.alignChildren = ["fill", "top"];
        var findReplaceTab = addDialogTab(modeTabbedPanel, LABELS.tab.findReplace);
        var englishTab = addDialogTab(modeTabbedPanel, LABELS.tab.englishText);
        var cleanupTab = addDialogTab(modeTabbedPanel, LABELS.tab.cleanup);
        modeTabbedPanel.selection = findReplaceTab;

        var findReplaceControls = buildFindReplacePanel(findReplaceTab, savedSettings, dialogState);
        var conversionPanelBuilder = createConversionPanelBuilder(onConvert);
        conversionPanelBuilder.addConversionPanel(englishTab, LABELS.panel.letterCase, ["caseUpper", "caseLower", "caseWord", "caseSentence", "caseTitle"]);
        buildCleanupPanels(cleanupTab, onConvert);

        /* 対象とオプションを2カラムに並べる / Place the scope and options panels in two columns */
        var scopeOptionsGroup = mainDialog.add("group");
        scopeOptionsGroup.orientation = "row";
        scopeOptionsGroup.alignChildren = ["fill", "fill"];
        var scopeRadios = buildScopePanel(scopeOptionsGroup, dialogState, hasSelection);
        buildOptionsPanel(scopeOptionsGroup, savedSettings, dialogState);

        return {
            mainDialog: mainDialog,
            findReplaceControls: findReplaceControls,
            conversionPanelBuilder: conversionPanelBuilder,
            scopeRadios: scopeRadios,
            buttonControls: buildButtonRow(mainDialog)
        };
    }

    /**
     * 入力内容とチェックボックス・対象の設定を保存する
     * @param {string[]} searchTexts - 検索欄の文字列
     * @param {string[]} replaceTexts - 置換欄の文字列
     * @param {Object} dialogState - showFindReplaceDialog() の状態
     * @returns {void}
     */
    function saveDialogSettings(searchTexts, replaceTexts, dialogState) {
        var settingsToSave = { searchTexts: searchTexts, replaceTexts: replaceTexts, scope: dialogState.scope };
        for (var i = 0; i < dialogState.settingControls.length; i++) {
            var settingKey = dialogState.settingControls[i].settingKey;
            settingsToSave[settingKey] = dialogState.values[settingKey];
        }
        saveSettings(settingsToSave);
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
     * @param {EditText[]} inputFields - 入力欄
     * @returns {string[]} 各欄の文字列
     */
    function readInputTexts(inputFields) {
        var inputTexts = [];
        for (var i = 0; i < inputFields.length; i++) {
            inputTexts.push(inputFields[i].text);
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
        setOptionalHelpTip(settingCheckbox, settingKey);
        dialogState.values[settingKey] = initialValue;
        dialogState.settingControls.push({ settingKey: settingKey, checkbox: settingCheckbox, defaultValue: defaultValue });
        settingCheckbox.onClick = function () {
            dialogState.values[settingKey] = this.value;
            if (dialogState.onSettingChange) dialogState.onSettingChange();
        };
        return settingCheckbox;
    }

    /**
     * LABELS.tooltip にキーがあれば、そのツールチップを付ける
     * @param {Object} control - 対象のコントロール
     * @param {string} tooltipKey - LABELS.tooltip のキー
     * @returns {void}
     */
    function setOptionalHelpTip(control, tooltipKey) {
        if (LABELS.tooltip[tooltipKey]) control.helpTip = getLabel(LABELS.tooltip[tooltipKey]);
    }

    /**
     * タブを追加する
     * @param {TabbedPanel} tabbedPanel - 追加先のタブパネル
     * @param {Object} labelSet - { ja, en } を持つタブ名
     * @returns {Tab} 追加したタブ
     */
    function addDialogTab(tabbedPanel, labelSet) {
        var dialogTab = tabbedPanel.add("tab", undefined, getLabel(labelSet));
        dialogTab.orientation = "column";
        dialogTab.alignChildren = ["fill", "top"];
        dialogTab.margins = TAB_MARGINS;
        return dialogTab;
    }

    /**
     * 「削除・置換する文字列」パネルを作る（入力欄・挿入ボタン・正規表現や空になったテキストの削除などのチェックボックス・リセット）
     * @param {Tab} parentTab - 追加先のタブ
     * @param {Object} savedSettings - loadSettings() の結果
     * @param {Object} dialogState - showFindReplaceDialog() の状態
     * @returns {Object} { searchInputs: EditText[], replaceInputs: EditText[], matchCountLabels: StaticText[], referenceButtonGroup: Group, btnReset: Button, focusFirstInput: Function }
     */
    function buildFindReplacePanel(parentTab, savedSettings, dialogState) {
        var savedSearchTexts = savedSettings.searchTexts || [];
        var savedReplaceTexts = savedSettings.replaceTexts || [];
        var findReplacePanel = parentTab.add("panel", undefined, getLabel(LABELS.panel.findReplace));
        setupPanel(findReplacePanel);

        var searchInputs = [];
        var replaceInputs = [];
        var matchCountLabels = [];
        /* 最後にカーソルがあった入力欄（挿入ボタンの挿入先）/ Field that last had the cursor, where the insert buttons insert */
        var lastActiveInput = null;
        /* ボタンを押すと入力欄のカーソルが失われるので、押し下げた時点の位置を控える / Clicking a button loses the caret, so keep its position from the mouse down */
        var savedCaret = null;

        /* 入力欄に、挿入先の記録とショートカットを付ける / Track the cursor field and add the shortcuts */
        function setupShortcutInput(targetInput) {
            targetInput.onActivate = function () {
                lastActiveInput = this;
            };
            targetInput.addEventListener("keydown", function (keyEvent) {
                var insertedToken = getShortcutToken(keyEvent.keyName, ScriptUI.environment.keyboardState, dialogState.values.useRegex);
                if (!insertedToken) return;
                /* Enter で OK が押されたり、option＋数字で記号が入ったりしないよう止める
                   Keep Enter from pressing OK and option+digit from typing a symbol */
                keyEvent.preventDefault();
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

        /* 左に正規表現などのチェックボックス、右下にリセット / Checkboxes on the left, Reset at the bottom right */
        var searchOptionsRowGroup = findReplacePanel.add("group");
        searchOptionsRowGroup.orientation = "row";
        searchOptionsRowGroup.alignment = ["fill", "top"];
        searchOptionsRowGroup.alignChildren = ["left", "bottom"];
        var searchOptionsGroup = searchOptionsRowGroup.add("group");
        searchOptionsGroup.orientation = "column";
        searchOptionsGroup.alignChildren = ["left", "top"];
        searchOptionsGroup.margins = [0, SEARCH_OPTIONS_TOP_MARGIN, 0, 0];
        searchOptionsGroup.spacing = PANEL_SPACING;
        addSettingCheckbox(searchOptionsGroup, "useRegex", DEFAULT_USE_REGEX, savedSettings, dialogState);
        addSettingCheckbox(searchOptionsGroup, "matchCase", DEFAULT_MATCH_CASE, savedSettings, dialogState);
        addSettingCheckbox(searchOptionsGroup, "deleteEmptyFrames", DEFAULT_DELETE_EMPTY_FRAMES, savedSettings, dialogState);
        var searchOptionsSpacer = searchOptionsRowGroup.add("group");
        searchOptionsSpacer.alignment = ["fill", "fill"];
        searchOptionsSpacer.minimumSize.width = 0;
        var btnReset = searchOptionsRowGroup.add("button", undefined, getLabel(LABELS.button.reset));
        btnReset.alignment = ["right", "bottom"];
        btnReset.helpTip = getLabel(LABELS.tooltip.reset);

        return {
            searchInputs: searchInputs,
            replaceInputs: replaceInputs,
            matchCountLabels: matchCountLabels,
            referenceButtonGroup: referenceButtonGroup,
            btnReset: btnReset,
            focusFirstInput: focusFirstInput
        };
    }

    /**
     * 変換ボタンと変換結果の見本を1行ずつ並べるパネルの作り手を用意する（見本はまとめて更新する）
     * @param {Function} onConvert - ボタンを押したときに (convertText, conversionKey) を受け取る処理
     * @returns {Object} { addConversionPanel: Function, updateSamples: Function }
     *   addConversionPanel(parentGroup, labelSet, conversionKeys) でパネルを追加し、updateSamples(sampleText) で見本を更新する
     */
    function createConversionPanelBuilder(onConvert) {
        /* 見本の一覧 / Sample rows */
        var sampleRows = [];

        /* ボタンを押したら変換関数を渡す / Pass the converter when the button is clicked */
        function setConvertHandler(conversionButton, conversionKey) {
            conversionButton.onClick = function () {
                onConvert(getTextConverter(conversionKey), conversionKey);
            };
        }

        return {
            addConversionPanel: function (parentGroup, labelSet, conversionKeys) {
                var conversionPanel = parentGroup.add("panel", undefined, getLabel(labelSet));
                setupPanel(conversionPanel);
                /* 見本をパネルの幅いっぱいに伸ばす / Stretch the samples to the panel width */
                conversionPanel.alignChildren = ["fill", "top"];
                for (var i = 0; i < conversionKeys.length; i++) {
                    var conversionRowGroup = addButtonRowGroup(conversionPanel);
                    var conversionButton = conversionRowGroup.add("button", undefined, getLabel(LABELS.button[conversionKeys[i]]));
                    conversionButton.preferredSize = CONVERSION_BUTTON_SIZE;
                    setOptionalHelpTip(conversionButton, conversionKeys[i]);
                    setConvertHandler(conversionButton, conversionKeys[i]);
                    var sampleLabel = conversionRowGroup.add("statictext", undefined, "");
                    sampleLabel.preferredSize = CONVERSION_SAMPLE_SIZE;
                    sampleLabel.alignment = ["fill", "center"];
                    sampleRows.push({ sampleLabel: sampleLabel, convertText: getTextConverter(conversionKeys[i]) });
                }
            },
            updateSamples: function (sampleText) {
                var shortSample = shortenSampleText(sampleText);
                for (var i = 0; i < sampleRows.length; i++) {
                    sampleRows[i].sampleLabel.text = shortenSampleText(sampleRows[i].convertText(shortSample));
                }
            }
        };
    }

    /**
     * 「整形」タブの縦のカラムを作る
     * @param {Group} parentGroup - 追加先のグループ
     * @returns {Group} 作ったカラム
     */
    function addCleanupColumn(parentGroup) {
        var cleanupColumnGroup = parentGroup.add("group");
        cleanupColumnGroup.orientation = "column";
        cleanupColumnGroup.alignment = ["fill", "top"];
        cleanupColumnGroup.alignChildren = ["fill", "top"];
        return cleanupColumnGroup;
    }

    /**
     * 「整形」タブの中身を作る（左にタブ・スペース、右に記号の変換・かな・数字の変換・リストの除去）
     * @param {Tab} parentTab - 追加先のタブ
     * @param {Function} onConvert - ボタンを押したときに (convertText, conversionKey) を受け取る処理
     * @returns {void}
     */
    function buildCleanupPanels(parentTab, onConvert) {
        var cleanupColumnsGroup = parentTab.add("group");
        cleanupColumnsGroup.orientation = "row";
        cleanupColumnsGroup.alignChildren = ["fill", "top"];
        var leftColumnGroup = addCleanupColumn(cleanupColumnsGroup);
        var rightColumnGroup = addCleanupColumn(cleanupColumnsGroup);

        /* 押すと変換を実行するボタンを作る / Create a button that runs a conversion */
        function addCleanupButton(parentGroup, labelKey) {
            var cleanupButton = parentGroup.add("button", undefined, getLabel(LABELS.button[labelKey]));
            setOptionalHelpTip(cleanupButton, labelKey);
            cleanupButton.onClick = function () {
                onConvert(getTextConverter(labelKey), labelKey);
            };
        }

        /* ボタンを縦に並べたパネルを作る / Create a panel of stacked buttons */
        function addCleanupPanel(parentColumn, labelSet, labelKeys) {
            var cleanupPanel = parentColumn.add("panel", undefined, getLabel(labelSet));
            setupPanel(cleanupPanel, CLEANUP_PANEL_MARGINS);
            cleanupPanel.alignChildren = ["fill", "top"];
            for (var i = 0; i < labelKeys.length; i++) {
                addCleanupButton(cleanupPanel, labelKeys[i]);
            }
            return cleanupPanel;
        }

        addCleanupPanel(leftColumnGroup, LABELS.panel.tabCharacter, ["removeTabs", "tabsToSpaces"]);
        addCleanupPanel(leftColumnGroup, LABELS.panel.removeSpace, ["trimSpaces", "cjkLatinSpaces", "collapseSpaces", "cleanupSpaces"]);
        addCleanupPanel(leftColumnGroup, LABELS.panel.addSpace, ["spaceAfterPunct"]);

        buildSymbolConversionPanel(rightColumnGroup, onConvert);
        /* かなと数字の変換は1つのパネルに2行で並べる / Put kana and digit conversions in one panel, on two rows */
        var kanaDigitPanel = addCleanupPanel(rightColumnGroup, LABELS.panel.kanaDigitConversion, []);
        kanaDigitPanel.alignChildren = ["left", "top"];
        var kanaButtonRowGroup = addButtonRowGroup(kanaDigitPanel);
        addCleanupButton(kanaButtonRowGroup, "toHiragana");
        addCleanupButton(kanaButtonRowGroup, "toKatakana");
        addCleanupButton(kanaButtonRowGroup, "toHalfKana");
        var digitButtonRowGroup = addButtonRowGroup(kanaDigitPanel);
        addCleanupButton(digitButtonRowGroup, "toHalfDigit");
        addCleanupButton(digitButtonRowGroup, "toFullDigit");

        /* リストの除去はボタンを横に並べ、伸ばさない / List removal buttons in a row, at their natural width */
        var removeListPanel = addCleanupPanel(rightColumnGroup, LABELS.panel.removeList, ["bulletList", "numberList"]);
        removeListPanel.orientation = "row";
        removeListPanel.alignChildren = ["left", "center"];
    }

    /**
     * 「スペースや記号の変換」パネルを作る（変換前・変換後をラジオで選び、［変換］で置き換える。変換前と同じ記号は変換後で選べない）
     * @param {Group} parentColumn - 追加先のカラム
     * @param {Function} onConvert - ［変換］を押したときに (convertText, conversionKey) を受け取る処理
     * @returns {void}
     */
    function buildSymbolConversionPanel(parentColumn, onConvert) {
        var symbolConversionPanel = parentColumn.add("panel", undefined, getLabel(LABELS.panel.symbolConversion));
        setupPanel(symbolConversionPanel, CLEANUP_PANEL_MARGINS);
        symbolConversionPanel.alignChildren = ["fill", "top"];
        /* 選んでいる記号。show() 前はラジオの value を読み戻せないので、ここに持つ
           Chosen symbols, kept here as radio values cannot be read back before show() */
        var symbolChoice = { before: "space", after: "underscore" };
        var afterRadios;

        /* 変換前で選んだ記号を変換後ではディムにし、重なったら変換後を空いている記号に移す
           Dim the Before symbol on the After side, and move After to a free symbol when they collide */
        function syncAfterRadios() {
            if (symbolChoice.after === symbolChoice.before) {
                symbolChoice.after = (symbolChoice.before === "space") ? "underscore" : "space";
            }
            for (var symbolKind in afterRadios) {
                afterRadios[symbolKind].enabled = (symbolKind !== symbolChoice.before);
                afterRadios[symbolKind].value = (symbolKind === symbolChoice.after);
            }
        }

        /* 変換前と変換後を2カラムに並べる / Place Before and After in two columns */
        var symbolColumnsGroup = symbolConversionPanel.add("group");
        symbolColumnsGroup.orientation = "row";
        symbolColumnsGroup.alignChildren = ["fill", "top"];
        addSymbolRadioPanel(symbolColumnsGroup, LABELS.panel.symbolBefore, symbolChoice, "before", syncAfterRadios);
        afterRadios = addSymbolRadioPanel(symbolColumnsGroup, LABELS.panel.symbolAfter, symbolChoice, "after", null);
        syncAfterRadios();

        var btnConvertSymbol = symbolConversionPanel.add("button", undefined, getLabel(LABELS.button.convertSymbol));
        btnConvertSymbol.alignment = ["center", "top"];
        setOptionalHelpTip(btnConvertSymbol, "convertSymbol");
        btnConvertSymbol.onClick = function () {
            onConvert(createSymbolConverter(symbolChoice.before, symbolChoice.after), "convertSymbol");
        };
    }

    /**
     * 記号（スペース・アンダースコア・ハイフン）を選ぶラジオのパネルを作る
     * @param {Group} parentGroup - 追加先のグループ
     * @param {Object} labelSet - { ja, en } を持つパネル名
     * @param {Object} symbolChoice - 選んでいる記号を持つオブジェクト
     * @param {string} choiceKey - symbolChoice のキー（"before" / "after"）
     * @param {Function|null} onChoose - 選び直したときの処理（無ければ null）
     * @returns {Object} 記号の名前をキーにしたラジオボタン { space, underscore, hyphen }
     */
    function addSymbolRadioPanel(parentGroup, labelSet, symbolChoice, choiceKey, onChoose) {
        var symbolRadioPanel = parentGroup.add("panel", undefined, getLabel(labelSet));
        setupPanel(symbolRadioPanel, CLEANUP_PANEL_MARGINS);
        var symbolKinds = ["space", "underscore", "hyphen"];
        var symbolRadios = {};

        /* クリックした記号を控える / Keep the clicked symbol */
        function addSymbolRadio(symbolKind) {
            var symbolRadio = symbolRadioPanel.add("radiobutton", undefined, getLabel(LABELS.radio[symbolKind]));
            symbolRadio.value = (symbolChoice[choiceKey] === symbolKind);
            symbolRadio.onClick = function () {
                symbolChoice[choiceKey] = symbolKind;
                if (onChoose) onChoose();
            };
            symbolRadios[symbolKind] = symbolRadio;
        }
        for (var i = 0; i < symbolKinds.length; i++) {
            addSymbolRadio(symbolKinds[i]);
        }
        return symbolRadios;
    }

    /**
     * 見本に表示するよう、改行や連続する空白を詰めて短くする
     * @param {string} text - 元の文字列
     * @returns {string} 1行に詰めた文字列（長ければ末尾を「…」に）
     */
    function shortenSampleText(text) {
        var sampleText = String(text).replace(/[\r\n\x03]+/g, " ").replace(/[ 　\t]+/g, " ").replace(/^\s+|\s+$/g, "");
        if (sampleText.length > CONVERSION_SAMPLE_MAX_CHARS) sampleText = sampleText.substring(0, CONVERSION_SAMPLE_MAX_CHARS) + "…";
        return sampleText;
    }

    /**
     * 選択している文字列を返す（変換結果の見本用）。文字ツールで選択した文字、無ければ選択したテキストの内容
     * @param {TextRange|null} selectedTextRange - 文字ツールで選択した文字
     * @param {PageItem[]} selectedItems - 実行時に選択していたアイテム
     * @returns {string} 選択している文字列。無ければ空文字列
     */
    function getSelectedSampleText(selectedTextRange, selectedItems) {
        if (selectedTextRange) {
            /* 変換で文字が消えると範囲が無効になり例外になる / The range throws once conversions remove its characters */
            try {
                if (selectedTextRange.contents !== "") return selectedTextRange.contents;
            } catch (e) {}
        }
        var selectedFrames = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i]) collectItemsOfType(selectedItems[i], "TextFrame", selectedFrames);
        }
        return getFirstNonEmptyText(getFrameContents(selectedFrames));
    }

    /**
     * 空でない最初の文字列を返す（変換結果の見本用）
     * @param {string[]} textContents - テキストの内容
     * @returns {string} 見つからなければ空文字列
     */
    function getFirstNonEmptyText(textContents) {
        for (var i = 0; i < textContents.length; i++) {
            if (textContents[i] !== "") return textContents[i];
        }
        return "";
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
        for (var scopeName in scopeRadios) {
            selectScopeOnClick(scopeName);
        }
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
        addSettingCheckbox(optionsPanel, "includeHidden", DEFAULT_INCLUDE_HIDDEN, savedSettings, dialogState);
        addSettingCheckbox(optionsPanel, "includeLocked", DEFAULT_INCLUDE_LOCKED, savedSettings, dialogState);
        addSettingCheckbox(optionsPanel, "includeSymbols", DEFAULT_INCLUDE_SYMBOLS, savedSettings, dialogState);
    }

    /**
     * ボタンエリアを作る（左にプレビュー、右にキャンセルと OK）
     * @param {Window} parentDialog - ダイアログ
     * @returns {Object} { previewCheckbox: Checkbox, btnOK: Button }
     */
    function buildButtonRow(parentDialog) {
        var btnRowGroup = parentDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var previewCheckbox = btnLeftGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.preview));
        previewCheckbox.helpTip = getLabel(LABELS.tooltip.preview);

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        return { previewCheckbox: previewCheckbox, btnOK: btnOK };
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
     * @param {EditText} targetInput - 入れる先の入力欄
     * @param {string} insertedToken - 入れる記号
     * @param {Object|null} caretPosition - captureCaret() の結果
     * @returns {void}
     */
    function insertTokenAtCaret(targetInput, insertedToken, caretPosition) {
        var currentText = targetInput.text;
        if (caretPosition && caretPosition.input === targetInput && caretPosition.start + caretPosition.length <= currentText.length) {
            targetInput.text = currentText.substring(0, caretPosition.start) + insertedToken + currentText.substring(caretPosition.start + caretPosition.length);
        } else {
            targetInput.text = currentText + insertedToken;
        }
    }

    /**
     * 入力欄のカーソル位置を読む。目印の文字を選択範囲に差し込んで位置を測り、元の文字列に戻す
     * （ScriptUI にはカーソル位置を返すプロパティが無いため）
     * @param {EditText} targetInput - 対象の入力欄（カーソルがあるうちに呼ぶ）
     * @returns {Object|null} { input: EditText, start: number, length: number }。読めなければ null
     */
    function captureCaret(targetInput) {
        var CARET_MARKER = "\u0001";
        var originalText = targetInput.text;
        var selectedLength = targetInput.textselection.length;
        targetInput.textselection = CARET_MARKER;
        var markerIndex = targetInput.text.indexOf(CARET_MARKER);
        targetInput.text = originalText;
        /* カーソルを失った入力欄では目印が先頭に入るので、先頭は信用しない（末尾に足す側に倒す）
           A field that lost its caret puts the marker at the start, so treat the start as unknown */
        if (markerIndex < 0 || (markerIndex === 0 && originalText.length > 0)) return null;
        return { input: targetInput, start: markerIndex, length: selectedLength };
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
        var patternSource = useRegex ? convertBreakTokensInRegex(searchText) : expandBreakTokens(searchText).replace(/[.*+?^${}()|[\]\\\/]/g, "\\$&");
        /* 誤った正規表現は new RegExp() が例外を出す / new RegExp() throws on an invalid pattern */
        try {
            return new RegExp(patternSource, matchCase ? "gm" : "gmi");
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
     * @param {string} regexSource - 入力された正規表現
     * @returns {string} 置き換えた正規表現
     */
    function convertBreakTokensInRegex(regexSource) {
        return regexSource.replace(/\\\\|\\n|@#/g, function (token) {
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
        for (var ancestorNode = item; ancestorNode && ancestorNode.typename !== "Document"; ancestorNode = ancestorNode.parent) {
            var isHidden = (ancestorNode.typename === "Layer") ? !ancestorNode.visible : ancestorNode.hidden;
            if (isHidden && !layerOptions.includeHidden) return false;
            if (ancestorNode.locked && !layerOptions.includeLocked) return false;
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
        for (var ancestorNode = item; ancestorNode && ancestorNode.typename !== "Document"; ancestorNode = ancestorNode.parent) {
            /* ロック中は表示を切り替えられないことがあるので、先にロックを外す / Unlock first, as a locked item may refuse to change visibility */
            if (ancestorNode.locked) {
                stateRecords.push({ target: ancestorNode, propertyName: "locked", originalValue: true });
                ancestorNode.locked = false;
            }
            if (ancestorNode.typename === "Layer") {
                if (!ancestorNode.visible) {
                    stateRecords.push({ target: ancestorNode, propertyName: "visible", originalValue: false });
                    ancestorNode.visible = true;
                }
            } else if (ancestorNode.hidden) {
                stateRecords.push({ target: ancestorNode, propertyName: "hidden", originalValue: true });
                ancestorNode.hidden = false;
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
     * 自身と親の非表示・ロックを一時的に解除して編集し、終わったら元に戻す
     * 編集できないテキスト（テンプレートレイヤーなど）は例外になるので、そのアイテムは飛ばす
     * @param {PageItem} item - 対象のアイテム
     * @param {Function} editItem - (stateRecords) を受け取る編集処理。消したアイテムの記録は stateRecords から除く
     * @returns {*} editItem の戻り値。例外のときは false
     */
    function editWithAncestorsReleased(item, editItem) {
        var stateRecords = [];
        try {
            unlockAndRevealAncestors(item, stateRecords);
            return editItem(stateRecords);
        } catch (e) {
            return false;
        } finally {
            restoreItemStates(stateRecords);
        }
    }

    /**
     * アイテム内から指定した型のアイテムを再帰的に集める（グループ内も含む）
     * @param {PageItem} item - 調べるアイテム
     * @param {string} typename - 集める型名
     * @param {PageItem[]} foundItems - 見つかったアイテムの追加先
     * @returns {void}
     */
    function collectItemsOfType(item, typename, foundItems) {
        if (item.typename === typename) {
            foundItems.push(item);
        } else if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) {
                collectItemsOfType(item.pageItems[i], typename, foundItems);
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
        var itemBounds = item.visibleBounds;
        return itemBounds[2] > artboardRect[0] &&
            itemBounds[0] < artboardRect[2] &&
            itemBounds[3] < artboardRect[1] &&
            itemBounds[1] > artboardRect[3];
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
            editWithAncestorsReleased(targetFrames[i], function (stateRecords) {
                var textFrame = targetFrames[i];
                if (!replacePatternsInFrame(textFrame, searchEntries, replaceResult.processedCounts)) return;
                replaceResult.changedCount++;
                if (deleteEmptyFrames && textFrame.contents === "" && isStandaloneFrame(textFrame)) {
                    /* 消したフレーム自身は状態を戻さない / Do not restore the deleted frame itself */
                    removeRecordsFor(stateRecords, textFrame);
                    textFrame.remove();
                    replaceResult.deletedFrameCount++;
                }
            });
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
     * 指定したオブジェクトの記録を取り除く（配列そのものを書き換える）
     * @param {Object[]} stateRecords - unlockAndRevealAncestors() の記録
     * @param {PageItem} removedItem - 記録を除くオブジェクト
     * @returns {void}
     */
    function removeRecordsFor(stateRecords, removedItem) {
        for (var i = stateRecords.length - 1; i >= 0; i--) {
            if (stateRecords[i].target === removedItem) stateRecords.splice(i, 1);
        }
    }

    /**
     * 文字列中の一致箇所を左から集める（長さ0の一致は除く）
     * @param {string} text - 検索する文字列
     * @param {RegExp} searchPattern - g フラグ付きの検索パターン
     * @returns {Object[]} { start: number, length: number, captures: Array } の配列（昇順）
     */
    function findMatches(text, searchPattern) {
        var matches = [];
        var regexMatch;
        searchPattern.lastIndex = 0;
        while ((regexMatch = searchPattern.exec(text)) !== null) {
            if (regexMatch[0].length === 0) {
                /* 長さ0の一致で止まらないよう1文字進める / Step past empty matches to avoid an endless loop */
                searchPattern.lastIndex++;
                continue;
            }
            matches.push({ start: regexMatch.index, length: regexMatch[0].length, captures: regexMatch });
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
    // 変換と整形 / Conversions and cleanup
    // =========================================
    /* 変換関数は TextBreakSplitMergePalette.jsx から移植 / Converters ported from TextBreakSplitMergePalette.jsx */

    /**
     * 変換の名前に対応する変換関数を返す
     * @param {string} conversionKey - LABELS.button のキー（caseUpper・removeTabs など）
     * @returns {Function|null} 文字列を受け取り、変換した文字列を返す関数。該当しなければ null
     */
    function getTextConverter(conversionKey) {
        switch (conversionKey) {
            case "caseUpper": return function (text) { return text.toUpperCase(); };
            case "caseLower": return function (text) { return text.toLowerCase(); };
            case "caseWord": return capitalizeEachWord;
            case "caseSentence": return toSentenceCase;
            case "caseTitle": return toTitleCase;
            /* かな変換は半角カナも受け付けるため、いったん全角カナへ寄せてから変換する / Kana conversions accept halfwidth kana, so normalize to fullwidth first */
            case "toHiragana": return function (text) { return toHiraganaText(toFullWidthKanaText(text)); };
            case "toKatakana": return function (text) { return toKatakanaText(toFullWidthKanaText(text)); };
            case "toHalfKana": return function (text) { return toHalfWidthKanaText(toKatakanaText(text)); };
            /* 数字変換は漢数字も受け付けるため、いったん算用数字へ寄せてから変換する / Digit conversions accept kanji numerals, so normalize to Arabic first */
            case "toHalfDigit": return function (text) { return toHalfWidthDigitText(toArabicNumeralText(text)); };
            case "toFullDigit": return function (text) { return toFullWidthDigitText(toArabicNumeralText(text)); };
            /* 整形 / Cleanup */
            case "removeTabs": return function (text) { return text.replace(/\t/g, ""); };
            case "tabsToSpaces": return function (text) { return text.replace(/\t/g, " "); };
            case "trimSpaces": return trimLineSpacesText;
            case "cjkLatinSpaces": return removeCjkLatinSpacesText;
            case "collapseSpaces": return collapseSpacesText;
            /* 先に連続を1つにまとめる（和欧間は前後の文字で判断するので、連続したままだと英単語間も消える）
               Collapse first; removing CJK/Latin spaces looks at neighbors, so runs of spaces would vanish between Latin words */
            case "cleanupSpaces": return function (text) { return removeCjkLatinSpacesText(collapseSpacesText(trimLineSpacesText(text))); };
            case "spaceAfterPunct": return function (text) { return text.replace(/([.,])(?=[^\s\d.,])/g, "$1 "); };
            case "bulletList": return removeBulletMarkersText;
            case "numberList": return removeNumberMarkersText;
        }
        return null;
    }

    /**
     * 単語の先頭のみ大文字。事前に小文字化しているので、すべて大文字の語も Negotiable のようになる
     * @param {string} text - 変換する文字列
     * @returns {string} 変換した文字列
     */
    function capitalizeEachWord(text) {
        return String(text).toLowerCase().replace(/\b([a-z])/g, function (matched, initial) {
            return initial.toUpperCase();
        });
    }

    /**
     * 文頭のみ大文字。文区切りが無い（英単語が1つだけの）場合も先頭を大文字化する
     * @param {string} text - 変換する文字列
     * @returns {string} 変換した文字列
     */
    function toSentenceCase(text) {
        return String(text).toLowerCase().replace(/(^|[\.\!\?]\s+|[\r\n]+)([a-z])/g,
            function (matched, prefix, initial) { return prefix + initial.toUpperCase(); });
    }

    /**
     * 英語タイトル形式（冠詞・前置詞などは小文字 / John Gruber の Title Caps 移植）
     * @param {string} text - 変換する文字列
     * @returns {string} 変換した文字列
     */
    function toTitleCase(text) {
        var smallWords = "(a|abaft|aboard|about|above|absent|across|afore|after|against|along|alongside|amid|amidst|among|amongst|an|and|apropos|around|as|aside|astride|at|athwart|atop|barring|before|behind|below|beneath|beside|besides|between|betwixt|beyond|but|by|circa|concerning|despite|down|during|except|excluding|failing|following|for|from|given|in|including|inside|into|lest|like|mid|midst|minus|modulo|near|next|nor|notwithstanding|of|off|on|onto|opposite|or|out|outside|over|pace|per|plus|pro|qua|regarding|round|sans|save|than|that|the|through|throughout|till|times|to|toward|towards|under|underneath|unlike|until|unto|up|upon|versus|via|vice|with|within|without|worth|v[.]?|via|vs[.]?)";
        var punctuation = "([!\"#$%&'()*+,./:;<=>?@[\\\\\\]^_`{|}~-]*)";

        function toLowerWord(word) { return word.toLowerCase(); }
        function capitalizeWord(word) { return word.substr(0, 1).toUpperCase() + word.substr(1); }

        var titleText = String(text);
        var sentenceSplitter = /[:.;?!] |(?: |^)[\"Ò]/g;
        var segments = [];
        var segmentStart = 0;
        while (true) {
            var splitMatch = sentenceSplitter.exec(titleText);
            segments.push(
                titleText.substring(segmentStart, splitMatch ? splitMatch.index : titleText.length)
                    .replace(/\b([A-Za-z][a-z.'Õ]*)\b/g, function (matched) {
                        return /[A-Za-z]\.[A-Za-z]/.test(matched) ? matched : capitalizeWord(matched);
                    })
                    .replace(RegExp("\\b" + smallWords + "\\b", "ig"), toLowerWord)
                    .replace(RegExp("^" + punctuation + smallWords + "\\b", "ig"), function (matched, leadingPunct, word) {
                        return leadingPunct + capitalizeWord(word);
                    })
                    .replace(RegExp("\\b" + smallWords + punctuation + "$", "ig"), capitalizeWord)
            );
            segmentStart = sentenceSplitter.lastIndex;
            if (!splitMatch) break;
            segments.push(splitMatch[0]);
        }
        return segments.join("")
            .replace(/ V(s?)\. /ig, " v$1. ")
            .replace(/(['Õ])S\b/ig, "$1s")
            .replace(/\b(AT&T|Q&A)\b/ig, function (matched) { return matched.toUpperCase(); });
    }

    /**
     * 半角カナを全角カナへ変換した文字列を返す（濁点・半濁点の合成と約物も対象）
     * @param {string} text - 変換する文字列
     * @returns {string} 変換した文字列
     */
    function toFullWidthKanaText(text) {
        var halfKana = "ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝ";
        var fullKana = "ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン";
        /* 濁点・半濁点の合成対応表 / Voiced and semi-voiced combinations */
        var dakutenBase = "ｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾊﾋﾌﾍﾎｳ";
        var dakutenFull = "ガギグゲゴザジズゼゾダヂヅデドバビブベボヴ";
        var handakutenBase = "ﾊﾋﾌﾍﾎ";
        var handakutenFull = "パピプペポ";
        /* 単独で置き換える約物（濁点・半濁点・句読点・カギ括弧・中黒）/ Punctuation replaced on its own */
        var halfPunct = "ﾞﾟ｡｢｣､･";
        var fullPunct = "゛゜。「」、・";

        var convertedText = "";
        for (var i = 0; i < text.length; i++) {
            var currentChar = text.charAt(i);
            var nextChar = text.charAt(i + 1);

            /* 直後が濁点・半濁点なら合成して1文字にする / Combine with a following voiced or semi-voiced mark */
            var isDakuten = (nextChar === "ﾞ");
            var comboIndex = -1;
            if (isDakuten) comboIndex = dakutenBase.indexOf(currentChar);
            else if (nextChar === "ﾟ") comboIndex = handakutenBase.indexOf(currentChar);
            if (comboIndex >= 0) {
                convertedText += isDakuten ? dakutenFull.charAt(comboIndex) : handakutenFull.charAt(comboIndex);
                i++;
                continue;
            }

            var kanaIndex = halfKana.indexOf(currentChar);
            var punctIndex = halfPunct.indexOf(currentChar);
            if (kanaIndex >= 0) {
                convertedText += fullKana.charAt(kanaIndex);
            } else if (punctIndex >= 0) {
                convertedText += fullPunct.charAt(punctIndex);
            } else {
                convertedText += currentChar;
            }
        }
        return convertedText;
    }

    /**
     * 全角カタカナをひらがなへ変換した文字列を返す（長音記号「ー」はそのまま）
     * ァ〜ヴと踊り字ヽヾは 0x60 引くとひらがなになる。ヵヶは「ヶ月」などの用例を壊すため対象外
     * @param {string} text - 変換する文字列
     * @returns {string} 変換した文字列
     */
    function toHiraganaText(text) {
        return text.replace(/[ァ-ヴヽヾ]/g, function (katakanaChar) {
            return String.fromCharCode(katakanaChar.charCodeAt(0) - 0x60);
        });
    }

    /**
     * ひらがなを全角カタカナへ変換した文字列を返す（長音記号「ー」はそのまま）
     * ぁ〜ゖと踊り字ゝゞは 0x60 足すとカタカナになる
     * @param {string} text - 変換する文字列
     * @returns {string} 変換した文字列
     */
    function toKatakanaText(text) {
        return text.replace(/[ぁ-ゖゝゞ]/g, function (hiraganaChar) {
            return String.fromCharCode(hiraganaChar.charCodeAt(0) + 0x60);
        });
    }

    /**
     * 全角カタカナを半角カナへ変換した文字列を返す（濁点・半濁点は2文字に分解、約物も対象）
     * @param {string} text - 変換する文字列
     * @returns {string} 変換した文字列
     */
    function toHalfWidthKanaText(text) {
        var fullKana = "ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン";
        var halfKana = "ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝ";
        /* 濁点・半濁点つきの文字を「素の半角カナ＋濁点」へ分解する対応表 / Split voiced characters into base kana and mark */
        var dakutenFull = "ガギグゲゴザジズゼゾダヂヅデドバビブベボヴ";
        var dakutenBase = "ｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾊﾋﾌﾍﾎｳ";
        var handakutenFull = "パピプペポ";
        var handakutenBase = "ﾊﾋﾌﾍﾎ";
        /* 単独で置き換える約物（濁点・半濁点・句読点・カギ括弧・中黒）/ Punctuation replaced on its own */
        var fullPunct = "゛゜。「」、・";
        var halfPunct = "ﾞﾟ｡｢｣､･";

        var convertedText = "";
        for (var i = 0; i < text.length; i++) {
            var currentChar = text.charAt(i);
            var dakutenIndex = dakutenFull.indexOf(currentChar);
            var handakutenIndex = handakutenFull.indexOf(currentChar);
            var kanaIndex = fullKana.indexOf(currentChar);
            var punctIndex = fullPunct.indexOf(currentChar);

            if (dakutenIndex >= 0) {
                convertedText += dakutenBase.charAt(dakutenIndex) + "ﾞ";
            } else if (handakutenIndex >= 0) {
                convertedText += handakutenBase.charAt(handakutenIndex) + "ﾟ";
            } else if (kanaIndex >= 0) {
                convertedText += halfKana.charAt(kanaIndex);
            } else if (punctIndex >= 0) {
                convertedText += halfPunct.charAt(punctIndex);
            } else {
                convertedText += currentChar;
            }
        }
        return convertedText;
    }

    /**
     * 全角数字を半角数字へ変換した文字列を返す
     * @param {string} text - 変換する文字列
     * @returns {string} 変換した文字列
     */
    function toHalfWidthDigitText(text) {
        return text.replace(/[０-９]/g, function (digitChar) {
            return String.fromCharCode(digitChar.charCodeAt(0) - 0xFEE0);
        });
    }

    /**
     * 半角数字を全角数字へ変換した文字列を返す
     * @param {string} text - 変換する文字列
     * @returns {string} 変換した文字列
     */
    function toFullWidthDigitText(text) {
        return text.replace(/[0-9]/g, function (digitChar) {
            return String.fromCharCode(digitChar.charCodeAt(0) + 0xFEE0);
        });
    }

    /**
     * 漢数字を算用数字（半角）へ変換した文字列を返す。位取りなしの表記（二〇二六）と、位取りありの表記（三十一・千二百三十四）の両方に対応する
     * @param {string} text - 変換する文字列
     * @returns {string} 変換した文字列
     */
    function toArabicNumeralText(text) {
        var kanjiDigits = "〇一二三四五六七八九";
        return text.replace(/[〇零一二三四五六七八九十百千万億]+/g, function (kanjiRun) {
            /* 単位の文字を含まないときは1文字ずつ置き換える（二〇二六 → 2026）/ Without unit characters, replace digit by digit */
            if (!/[十百千万億]/.test(kanjiRun)) {
                return kanjiRun.replace(/./g, function (kanjiChar) {
                    return (kanjiChar === "零") ? "0" : String(kanjiDigits.indexOf(kanjiChar));
                });
            }

            /* 位取りありは、万・億でいったん確定させながら積み上げる（二万五千 → 25000）/ Accumulate, settling at 万 and 億 */
            var smallUnitValues = { "十": 10, "百": 100, "千": 1000 };
            var largeUnitValues = { "万": 10000, "億": 100000000 };
            var totalValue = 0;
            var sectionValue = 0;
            var currentDigit = 0;
            for (var i = 0; i < kanjiRun.length; i++) {
                var kanjiChar = kanjiRun.charAt(i);
                if (smallUnitValues[kanjiChar]) {
                    sectionValue += (currentDigit || 1) * smallUnitValues[kanjiChar];
                    currentDigit = 0;
                } else if (largeUnitValues[kanjiChar]) {
                    totalValue += (sectionValue + currentDigit) * largeUnitValues[kanjiChar];
                    sectionValue = 0;
                    currentDigit = 0;
                } else {
                    /* 零は 0、それ以外は漢数字の値 / 零 is 0, others are their digit values */
                    currentDigit = (kanjiChar === "零") ? 0 : kanjiDigits.indexOf(kanjiChar);
                }
            }
            return String(totalValue + sectionValue + currentDigit);
        });
    }

    /**
     * 記号（スペース・アンダースコア・ハイフン）を別の記号に置き換える変換関数を作る
     * @param {string} beforeSymbol - 変換前（"space" / "underscore" / "hyphen"）。スペースは半角・全角の両方
     * @param {string} afterSymbol - 変換後（"space" / "underscore" / "hyphen"）
     * @returns {Function} 文字列を受け取り、変換した文字列を返す関数
     */
    function createSymbolConverter(beforeSymbol, afterSymbol) {
        var symbolPatterns = { space: /[ 　]/g, underscore: /_/g, hyphen: /-/g };
        var symbolTexts = { space: " ", underscore: "_", hyphen: "-" };
        return function (text) {
            return text.replace(symbolPatterns[beforeSymbol], symbolTexts[afterSymbol]);
        };
    }

    /**
     * 各段落の行頭・行末のスペースとタブを削除する
     * @param {string} text - 変換する文字列
     * @returns {string} 変換した文字列
     */
    function trimLineSpacesText(text) {
        var lines = text.split("\r");
        for (var i = 0; i < lines.length; i++) {
            lines[i] = lines[i].replace(/^[ \t　]+/, "").replace(/[ \t　]+$/, "");
        }
        return lines.join("\r");
    }

    /**
     * 連続する半角スペース・全角スペースを1つにまとめる
     * @param {string} text - 変換する文字列
     * @returns {string} 変換した文字列
     */
    function collapseSpacesText(text) {
        return text.replace(/ {2,}/g, " ").replace(/　{2,}/g, "　");
    }

    /**
     * 和欧間のスペースを削除する（前後がどちらも英数字のスペースは残す）
     * @param {string} text - 変換する文字列
     * @returns {string} 変換した文字列
     */
    function removeCjkLatinSpacesText(text) {
        return text.replace(/[ 　]/g, function (spaceChar, spaceIndex) {
            return (isLatinLetterOrDigit(text.charAt(spaceIndex - 1)) && isLatinLetterOrDigit(text.charAt(spaceIndex + 1))) ? spaceChar : "";
        });
    }

    /**
     * 半角の英字か数字か判定する
     * @param {string} character - 判定する1文字（空文字列なら false）
     * @returns {boolean} 半角の英数字なら true
     */
    function isLatinLetterOrDigit(character) {
        return /^[A-Za-z0-9]$/.test(character);
    }

    /**
     * 各段落の行頭から、最初に一致したパターンを1つだけ取り除く
     * @param {string} text - 変換する文字列
     * @param {RegExp[]} prefixPatterns - 行頭に一致させるパターン（先に書いたものを優先）
     * @returns {string} 変換した文字列
     */
    function removeLinePrefixText(text, prefixPatterns) {
        var lines = text.split("\r");
        for (var i = 0; i < lines.length; i++) {
            for (var j = 0; j < prefixPatterns.length; j++) {
                if (!prefixPatterns[j].test(lines[i])) continue;
                lines[i] = lines[i].replace(prefixPatterns[j], "");
                break;
            }
        }
        return lines.join("\r");
    }

    /**
     * 行頭の箇条書き記号を取り除く（「タブ＋記号＋タブ」の形と手入力の両方。「-」「*」は直後が空白のときだけ）
     * @param {string} text - 変換する文字列
     * @returns {string} 変換した文字列
     */
    function removeBulletMarkersText(text) {
        return removeLinePrefixText(text, [
            /^\t[・･·•◦●○◎□■▪◆◇✓–—\-\*]\t/,
            /^[\t 　]*(?:[・･·•◦●○◎□■▪◆◇✓]|[–—\-\*](?=[\t 　]))[\t 　]*/
        ]);
    }

    /**
     * 行頭の番号を取り除く（数字・全角数字・丸数字・英字・漢数字。区切りは . ． : ： |）
     * 「12.5」のように区切りの直後が数字なら本文とみなして残す
     * @param {string} text - 変換する文字列
     * @returns {string} 変換した文字列
     */
    function removeNumberMarkersText(text) {
        return removeLinePrefixText(text, [
            /^\t(?:[①-⑳❶-❿⓫-⓴]|[A-Za-z]+|[〇一二三四五六七八九十百千]+|[0-9０-９]+)[.．:：|]?\t/,
            /^[\t 　]*[①-⑳❶-❿⓫-⓴][\t 　]*/,
            /^[\t 　]*(?:[A-Za-z]+|[〇一二三四五六七八九十百千]+|[0-9０-９]+)[.．:：|][\t 　]+/,
            /^[\t 　]*[0-9０-９]+[.．](?![0-9０-９])[\t 　]*/
        ]);
    }

    /**
     * テキストフレームを変換する（非表示・ロックは処理のあいだだけ解除する）
     * @param {TextFrame[]} targetFrames - 対象のテキストフレーム
     * @param {Function} convertText - getTextConverter() の結果
     * @param {Function|null} prepareFrame - 変換の前に (textFrame) を受け取って行う処理（箇条書きのテキスト化など）。無ければ null
     * @returns {number} 変更したテキストフレームの数
     */
    function convertTextInFrames(targetFrames, convertText, prepareFrame) {
        var changedCount = 0;
        for (var i = 0; i < targetFrames.length; i++) {
            var isChanged = editWithAncestorsReleased(targetFrames[i], function () {
                if (prepareFrame) prepareFrame(targetFrames[i]);
                return convertFrameKeepingFormat(targetFrames[i], convertText);
            });
            if (isChanged) changedCount++;
        }
        return changedCount;
    }

    /**
     * テキストフレームの内容を変換し、変わった文字だけを書き換える（contents を書き戻さないので文字ごとの書式が残る）
     * 長さが変わらない変換は1文字ずつ、文字を削るだけ・足すだけの変換（スペースの削除・追加など）は変わる文字だけ、
     * それ以外（濁点の合成・漢数字など）は空白や改行で区切った語ごとに差し替える
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Function} convertText - getTextConverter() の結果
     * @returns {boolean} 書き換えたら true
     */
    function convertFrameKeepingFormat(textFrame, convertText) {
        var originalText = textFrame.contents;
        var convertedText = convertText(originalText);
        if (convertedText === originalText || originalText === "") return false;
        var frameCharacters = textFrame.characters;
        if (convertedText.length === originalText.length) {
            replaceDifferingPart(frameCharacters, 0, originalText, convertedText);
            return true;
        }
        var isShortened = convertedText.length < originalText.length;
        var matchedIndexes = isShortened ? matchSubsequence(convertedText, originalText) : matchSubsequence(originalText, convertedText);
        if (matchedIndexes) {
            if (isShortened) {
                removeUnkeptCharacters(frameCharacters, originalText.length, matchedIndexes);
            } else {
                insertAddedText(frameCharacters, originalText, convertedText, matchedIndexes);
            }
            return true;
        }
        /* 長さが変わる変換は前後の文字に左右されないので、語ごとに変換して右から差し替える
           Length-changing conversions do not depend on context, so convert word by word from the right */
        var wordMatches = findMatches(originalText, /[^\s\x03]+/g);
        for (var i = wordMatches.length - 1; i >= 0; i--) {
            var originalWord = wordMatches[i].captures[0];
            var convertedWord = convertText(originalWord);
            if (convertedWord !== originalWord) replaceDifferingPart(frameCharacters, wordMatches[i].start, originalWord, convertedWord);
        }
        return true;
    }

    /**
     * 短い文字列の各文字が、長い文字列のどこに対応するかを左から探す（長い文字列から文字を抜くだけで短い文字列になるか）
     * @param {string} shortText - 短い文字列
     * @param {string} longText - 長い文字列
     * @returns {number[]|null} shortText の各文字に対応する longText の位置。抜くだけでは作れなければ null
     */
    function matchSubsequence(shortText, longText) {
        var matchedIndexes = [];
        var longIndex = 0;
        for (var i = 0; i < shortText.length; i++) {
            while (longIndex < longText.length && longText.charAt(longIndex) !== shortText.charAt(i)) longIndex++;
            if (longIndex >= longText.length) return null;
            matchedIndexes.push(longIndex);
            longIndex++;
        }
        return matchedIndexes;
    }

    /**
     * 残す文字以外を右から削除する
     * @param {TextFrameItem[]} frameCharacters - textFrame.characters
     * @param {number} originalLength - 変換前の文字数
     * @param {number[]} keptIndexes - 残す文字の位置（昇順）
     * @returns {void}
     */
    function removeUnkeptCharacters(frameCharacters, originalLength, keptIndexes) {
        var keptIndex = keptIndexes.length - 1;
        for (var i = originalLength - 1; i >= 0; i--) {
            if (keptIndex >= 0 && keptIndexes[keptIndex] === i) {
                keptIndex--;
                continue;
            }
            frameCharacters[i].remove();
        }
    }

    /**
     * 増えた文字を、右から隣の文字に足して入れる（足した文字はその隣の文字の書式になる）
     * @param {TextFrameItem[]} frameCharacters - textFrame.characters
     * @param {string} originalText - 変換前の文字列（空でないこと）
     * @param {string} convertedText - 変換後の文字列
     * @param {number[]} originalIndexes - 変換前の各文字に対応する変換後の位置
     * @returns {void}
     */
    function insertAddedText(frameCharacters, originalText, convertedText, originalIndexes) {
        for (var i = originalText.length - 1; i >= 0; i--) {
            var nextIndex = (i + 1 < originalText.length) ? originalIndexes[i + 1] : convertedText.length;
            var addedAfter = convertedText.substring(originalIndexes[i] + 1, nextIndex);
            var addedBefore = (i === 0) ? convertedText.substring(0, originalIndexes[0]) : "";
            if (addedAfter !== "" || addedBefore !== "") frameCharacters[i].contents = addedBefore + originalText.charAt(i) + addedAfter;
        }
    }

    /**
     * 変換前と変換後で異なる部分だけを書き換える
     * 長さが同じなら異なる文字を1文字ずつ、違えば前後の一致部分を除いた範囲を差し替える（差し替えた文字はその範囲の先頭の文字の書式になる）
     * @param {TextFrameItem[]} frameCharacters - textFrame.characters
     * @param {number} offset - 書き換える範囲の開始位置
     * @param {string} originalText - 変換前の文字列
     * @param {string} convertedText - 変換後の文字列
     * @returns {void}
     */
    function replaceDifferingPart(frameCharacters, offset, originalText, convertedText) {
        var i;
        if (convertedText.length === originalText.length) {
            for (i = originalText.length - 1; i >= 0; i--) {
                if (convertedText.charAt(i) !== originalText.charAt(i)) frameCharacters[offset + i].contents = convertedText.charAt(i);
            }
            return;
        }
        var shorterLength = Math.min(originalText.length, convertedText.length);
        var prefixLength = 0;
        while (prefixLength < shorterLength && originalText.charAt(prefixLength) === convertedText.charAt(prefixLength)) prefixLength++;
        var suffixLength = 0;
        while (suffixLength < shorterLength - prefixLength &&
            originalText.charAt(originalText.length - 1 - suffixLength) === convertedText.charAt(convertedText.length - 1 - suffixLength)) suffixLength++;

        var replaceStart = offset + prefixLength;
        var replaceEnd = offset + originalText.length - suffixLength;
        var middleText = convertedText.substring(prefixLength, convertedText.length - suffixLength);
        if (replaceStart === replaceEnd) {
            /* 文字が増えるだけのときは、隣の文字に足す / When characters are only added, attach them to a neighbor */
            if (prefixLength > 0) {
                frameCharacters[replaceStart - 1].contents = originalText.charAt(prefixLength - 1) + middleText;
            } else {
                frameCharacters[replaceStart].contents = middleText + originalText.charAt(prefixLength);
            }
            return;
        }
        var keptCount = (middleText === "") ? 0 : 1;
        for (i = replaceEnd - 1; i >= replaceStart + keptCount; i--) {
            frameCharacters[i].remove();
        }
        if (keptCount > 0) frameCharacters[replaceStart].contents = middleText;
    }

    /**
     * テキストの箇条書き・番号付きリストを［テキストに変換］で本文の記号にする（行頭が「•＋タブ」などになる）
     * listStyle はスクリプトから読めず「なし」も書けないので、選択してメニューのコマンドを実行する
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function convertListStyleToText(doc, textFrame) {
        doc.selection = null;
        textFrame.selected = true;
        /* 選択を変えた直後は古い選択のままコマンドが走ることがあるので、先に再描画する
           Redraw first, as a command right after changing the selection may still see the old one */
        app.redraw();
        app.executeMenuCommand("convert list style to text");
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
            var sourceSymbol = symbolItems[i].symbol;
            if (indexOfItem(targetSymbols, sourceSymbol) === -1) targetSymbols.push(sourceSymbol);
        }
        return targetSymbols;
    }

    /**
     * 配列の中でのオブジェクトの位置を返す（DOM の参照は === で比べられる）
     * @param {Object[]} itemList - 探す配列
     * @param {Object} searchedItem - 探すオブジェクト
     * @returns {number} 位置。無ければ -1
     */
    function indexOfItem(itemList, searchedItem) {
        for (var i = 0; i < itemList.length; i++) {
            if (itemList[i] === searchedItem) return i;
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
        forEachSymbolContent(doc, targetSymbols, function (targetSymbol, contentGroup, symbolFrames) {
            textContents = textContents.concat(getFrameContents(symbolFrames));
        });
        return textContents;
    }

    /**
     * シンボルを1つずつ作業レイヤーに展開して処理し、展開したアートは処理のあとに消す
     * @param {Document} doc - 対象ドキュメント
     * @param {Symbol[]} targetSymbols - 対象のシンボル
     * @param {Function} handleSymbol - (targetSymbol, contentGroup, symbolFrames) を受け取る処理
     * @returns {void}
     */
    function forEachSymbolContent(doc, targetSymbols, handleSymbol) {
        if (targetSymbols.length === 0) return;
        withSymbolWorkLayer(doc, function (workLayer) {
            for (var i = 0; i < targetSymbols.length; i++) {
                var contentGroup = expandSymbolToGroup(doc, targetSymbols[i], workLayer);
                var symbolFrames = [];
                collectItemsOfType(contentGroup, "TextFrame", symbolFrames);
                handleSymbol(targetSymbols[i], contentGroup, symbolFrames);
                contentGroup.remove();
            }
        });
    }

    /**
     * シンボル内のテキストを書き換え、変わったシンボルだけ定義を差し替える
     * @param {Document} doc - 対象ドキュメント
     * @param {Symbol[]} targetSymbols - 対象のシンボル
     * @param {Function} editSymbolFrames - (symbolFrames) を受け取って書き換え、変えたら差し替えの成功後に呼ぶ関数を、変えなければ null を返す
     * @returns {number} 書き換えたシンボルの数
     */
    function rewriteSymbols(doc, targetSymbols, editSymbolFrames) {
        var updatedSymbolCount = 0;
        forEachSymbolContent(doc, targetSymbols, function (targetSymbol, contentGroup, symbolFrames) {
            var onSymbolReplaced = editSymbolFrames(symbolFrames);
            if (onSymbolReplaced && replaceSymbolDefinition(doc, targetSymbol, contentGroup)) {
                updatedSymbolCount++;
                onSymbolReplaced();
            }
        });
        return updatedSymbolCount;
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
        return rewriteSymbols(doc, targetSymbols, function (symbolFrames) {
            /* 差し替えに失敗したら数えないよう、シンボルごとに数えてから足す / Count per symbol, and add only when the swap succeeds */
            var symbolProcessedCounts = createZeroCounts(searchEntries.length);
            var isChanged = false;
            for (var j = 0; j < symbolFrames.length; j++) {
                if (replacePatternsInFrame(symbolFrames[j], searchEntries, symbolProcessedCounts)) isChanged = true;
            }
            if (!isChanged) return null;
            return function () {
                for (var k = 0; k < searchEntries.length; k++) processedCounts[k] += symbolProcessedCounts[k];
            };
        });
    }

    /**
     * シンボル内のテキストを変換し、書き換えたシンボルで定義を差し替える
     * @param {Document} doc - 対象ドキュメント
     * @param {Symbol[]} targetSymbols - 対象のシンボル
     * @param {Function} convertText - getTextConverter() の結果
     * @param {Function|null} prepareFrame - 変換の前に (textFrame) を受け取って行う処理。無ければ null
     * @returns {number} 書き換えたシンボルの数
     */
    function convertTextInSymbols(doc, targetSymbols, convertText, prepareFrame) {
        return rewriteSymbols(doc, targetSymbols, function (symbolFrames) {
            var isChanged = false;
            for (var j = 0; j < symbolFrames.length; j++) {
                if (prepareFrame) prepareFrame(symbolFrames[j]);
                if (convertFrameKeepingFormat(symbolFrames[j], convertText)) isChanged = true;
            }
            /* 変換では数えるものが無いので、差し替え後の処理は空 / Nothing to count after a conversion */
            return isChanged ? function () {} : null;
        });
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
