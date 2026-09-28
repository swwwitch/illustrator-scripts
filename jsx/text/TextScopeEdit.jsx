#target illustrator
#targetengine "TextScopeEditEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内のテキスト（シンボル内を含む）を一覧にし、書式を保ったまま編集して書き戻します。
対象はアートボード単位やレイヤー単位で絞り込め、テキストとフォント名の書き出しもできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextScopeEdit.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nb845889dd553

### Overview

Lists the text in the document, including text in symbols, and writes your edits back while keeping the formatting.
The scope can be narrowed by artboard or layer, and the text and font names can be exported.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextScopeEdit.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextScopeEdit";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-08";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextScopeEdit.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextScopeEdit.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nb845889dd553"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var SOFT_BREAK              = '@#';  /* 編集欄での強制改行の代替文字 / Placeholder for forced line breaks in the edit field */
    var POSITION_SORT_TOLERANCE = 10;    /* 位置順で同じ行とみなす高さの差（pt）/ Height difference treated as the same row (pt) */
    var LIST_LABEL_MAX_LENGTH   = 40;    /* 一覧に出す文字数 / Characters shown per list row */
    var DEFAULT_INCLUDE_SYMBOLS = true;  /* ［シンボル内のテキスト］の初期値 / Initial value of Text in Symbols */

    // =========================================
    // 作業レイヤー / Work layer
    // =========================================

    /* シンボルを展開するための作業レイヤー名 / Name of the work layer used to expand symbols */
    var WORK_LAYER_NAME = '__TextScopeEdit_work__';

    // =========================================
    // レイアウト / Layout
    // =========================================

    var TABS_SIZE                = [550, 350];         /* タブ全体の大きさ / Tabbed panel size */
    var TABS_MARGINS             = [15, 10, 1, 10];    /* タブ全体の余白 [左,上,右,下] / Tabbed panel margins */
    var EDIT_TAB_MARGINS         = [15, 20, 1, 10];    /* ［テキスト］タブの余白 / Text tab margins */
    var EDIT_TAB_SPACING         = 15;                 /* ［テキスト］タブの要素間隔 / Text tab spacing */
    var INFO_TAB_MARGINS         = [15, 20, 0, 10];    /* 名前一覧タブの余白 / Name-list tab margins */
    var INFO_TAB_SPACING         = 10;                 /* 名前一覧タブの要素間隔 / Name-list tab spacing */
    var CHOICE_ROW_SPACING       = 15;                 /* ラジオ・チェックボックスの横並びの間隔 / Spacing of radio/checkbox rows */
    var SCOPE_PANEL_MARGINS      = [15, 20, 1, 10];    /* 右カラムのパネル余白 / Right-column panel margins */
    var EXPORT_PANEL_MARGINS     = [15, 20, 15, 10];   /* 書き出しオプションのパネル余白 / Export options panel margins */
    var TEXT_LIST_BOUNDS         = [0, 0, 250, 270];   /* テキスト一覧の大きさ（［更新］ボタンの分を引く）/ Text list bounds (minus the Update button) */
    var TEXT_EDIT_BOUNDS         = [0, 0, 250, 72];    /* テキスト編集欄の大きさ / Text edit field bounds */
    var FONT_COLUMN_WIDTHS       = [180, 180, 120];    /* フォント一覧の列幅 / Font list column widths */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    var LABELS = {
        dialog: {
            title: { ja: "テキストの収集と編集", en: "Collect and Edit Text" },
            exportOptions: { ja: "書き出しオプション", en: "Export Options" }
        },
        tab: {
            editText: { ja: "テキスト", en: "Text" },
            layerNames: { ja: "レイヤー名", en: "Layer Names" },
            artboardNames: { ja: "アートボード名", en: "Artboard Names" },
            fontNames: { ja: "フォント名", en: "Font Names" }
        },
        panel: {
            artboardScope: { ja: "対象範囲", en: "Scope" },
            targetText: { ja: "対象テキスト", en: "Text to Include" },
            sort: { ja: "並び順", en: "Sort" },
            editOptions: { ja: "オプション", en: "Options" },
            exportContent: { ja: "書き出す内容", en: "Contents to Export" }
        },
        radio: {
            currentArtboard: { ja: "現在のアートボード", en: "Current Artboard" },
            allArtboards: { ja: "すべてのアートボード", en: "All Artboards" },
            sortNone: { ja: "なし", en: "None" },
            sortPosition: { ja: "位置順", en: "By Position" },
            sortAlphabetical: { ja: "ABC順", en: "Alphabetical" },
            layerScopeAll: { ja: "すべてのレイヤー", en: "All Layers" },
            layerScopeTop: { ja: "上位レベルのレイヤーのみ", en: "Top-Level Layers Only" },
            artboardNameNumbered: { ja: "番号つき", en: "With Numbers" },
            artboardNameOnly: { ja: "アートボード名のみ", en: "Names Only" }
        },
        checkbox: {
            mergeDuplicates: { ja: "同じ内容を一括編集", en: "Edit Identical Text Together" },
            selectedOnly: { ja: "選択しているテキストのみ", en: "Selected Text Only" },
            wholeDocument: { ja: "アートボード外も対象", en: "Include Outside Artboards" },
            includeCommentLayers: { ja: "「//」ではじまるレイヤー", en: "Layers Starting with //" },
            includeLocked: { ja: "ロックされたテキスト", en: "Locked Text" },
            includeHidden: { ja: "非表示のテキスト", en: "Hidden Text" },
            includeSymbols: { ja: "シンボル内のテキスト", en: "Text in Symbols" },
            keepFormat: { ja: "書式を保持", en: "Keep Formatting" },
            exportIncludeText: { ja: "テキスト", en: "Text" },
            exportIncludeFonts: { ja: "フォント名", en: "Font Names" },
            exportOpenAfter: { ja: "書き出し後にファイルを開く", en: "Open File After Export" },
            fontPS: { ja: "PostScript名", en: "PostScript Name" },
            fontFamily: { ja: "フォント名", en: "Font Name" },
            fontStyle: { ja: "スタイル", en: "Style" }
        },
        fieldLabel: {
            textList: { ja: "テキスト一覧", en: "Text List" },
            textEdit: { ja: "テキスト編集", en: "Edit Text" },
            layerNameList: { ja: "レイヤー名一覧", en: "Layer Name List" },
            artboardNameList: { ja: "アートボード名一覧", en: "Artboard Name List" },
            fontList: { ja: "フォント一覧", en: "Font List" }
        },
        listColumn: {
            fontPS: { ja: "PostScript名", en: "PostScript Name" },
            fontFamily: { ja: "フォント名", en: "Font Name" },
            fontStyle: { ja: "スタイル", en: "Style" }
        },
        format: {
            itemPrefix: { ja: ": ", en: ": " },
            symbolSuffix: { ja: "〈シンボル：{symbolName}〉", en: " «Symbol: {symbolName}»" },
            symbolRowMark: { ja: "♣ ", en: "♣ " },
            symbolRowSuffix: { ja: "〈{symbolName}〉", en: " «{symbolName}»" }
        },
        tooltip: {
            keepFormat: {
                ja: "変更した文字だけを書き換え、文字と段落の書式を保持します\n（オフにすると全体を1文字目の書式で置き換えます）",
                en: "Rewrites only the changed characters, keeping character\nand paragraph formatting (when off, the whole text\ntakes the formatting of its first character)"
            },
            wholeDocument: {
                ja: "アートボードの外にあるテキストも対象にします\n（［すべてのアートボード］のときに選べます）",
                en: "Also includes text outside the artboards\n(available with All Artboards)"
            },
            mergeDuplicates: {
                ja: "同じ内容のテキストを一覧の1行にまとめ、\n編集をそのすべてに反映します",
                en: "Lists identical text as one row and applies\nthe edit to every copy"
            },
            selectedOnly: {
                ja: "編集を、起動時に選択していたテキストにだけ反映します\n（シンボル内のテキストは、インスタンスを選択していたときだけ）",
                en: "Applies the edit only to the text selected when the script started\n(text in a symbol only when one of its instances was selected)"
            },
            sortPosition: {
                ja: "上から下へ並べ、ほぼ同じ高さのものは\n左から右へ並べます",
                en: "Sorts top to bottom, and left to right\nfor text at about the same height"
            },
            sortAlphabetical: {
                ja: "大文字と小文字を区別せずに\n内容の順に並べます",
                en: "Sorts by content, ignoring case"
            },
            textEdit: {
                ja: "Shift+Enter で強制改行を入れます\n（{softBreak} と表示されます）",
                en: "Shift+Enter inserts a forced line break\n(shown as {softBreak})"
            },
            textList: {
                ja: "♣ の行（シンボル内のテキスト）を編集すると、OK でシンボルの定義を書き換えます。\n範囲外にある同じシンボルのインスタンスも変わり、\n基準点などのシンボルオプションは初期値になります",
                en: "Editing a ♣ row (text in a symbol) rewrites the symbol definition on OK,\nso instances outside the scope change too, and symbol options\nsuch as the registration point are reset"
            },
            copyText: {
                ja: "テキスト一覧に並んでいるテキストを\n省略せずにクリップボードへコピーします",
                en: "Copies the full text of every row\nin the text list to the clipboard"
            },
            exportText: {
                ja: "アートボードごとにテキストとフォント名をまとめ、\nデスクトップにテキストファイルで書き出します",
                en: "Writes the text and font names, grouped by artboard,\nto a text file on the desktop"
            },
            fontList: {
                ja: "行をクリックすると、そのフォントを使っている\nテキストを選択します（シンボルはインスタンスを選択）",
                en: "Click a row to select the text that uses that font\n(for symbols, their instances are selected)"
            },
            includeCommentLayers: {
                ja: "名前が「//」ではじまるレイヤー（メモ用など）の\nテキストも対象にします",
                en: "Also includes text on layers whose names\nstart with // (such as note layers)"
            },
            includeLocked: {
                ja: "ロックされたテキストと、ロックされた\nレイヤー・グループ内のテキストも対象にします",
                en: "Also includes locked text and text\nin locked layers or groups"
            },
            includeHidden: {
                ja: "非表示のテキストと、非表示の\nレイヤー・グループ内のテキストも対象にします",
                en: "Also includes hidden text and text\nin hidden layers or groups"
            },
            includeSymbols: {
                ja: "対象範囲にあるシンボルのテキストを\nテキスト一覧の末尾に並べ、書き出しにも含めます",
                en: "Lists the text in symbols within the scope at the end\nof the text list and includes it in the export"
            },
            updateText: {
                ja: "編集をドキュメントに反映し、一覧を更新します\n（ダイアログは閉じません。キャンセルしても元に戻りません）",
                en: "Applies the edit to the document and refreshes the list\n(the dialog stays open; Cancel does not revert it)"
            }
        },
        button: {
            exportText: { ja: "テキスト書き出し...", en: "Export Text..." },
            copyText: { ja: "テキストをコピー", en: "Copy Text" },
            updateText: { ja: "更新", en: "Update" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        fallbackName: {
            artboardNumber: { ja: "アートボード{number}", en: "Artboard {number}" },
            outsideArtboards: { ja: "アートボード外", en: "Outside Artboards" },
            unknownFont: { ja: "不明", en: "Unknown" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません", en: "No document is open" },
            exportFailed: { ja: "テキストを書き出せませんでした", en: "Failed to export text" },
            noMatchingFontText: { ja: "該当するテキストがありません", en: "No matching text found" },
            symbolUpdateFailed: { ja: "シンボルを書き換えられませんでした", en: "Failed to update the symbol" }
        }
    };

    // =========================================
    // 文字列の整形 / String helpers
    // =========================================

    /**
     * 一覧用のラベルを作る（改行を空白にし、長ければ先頭だけにする）
     * @param {string} text - テキストの内容
     * @returns {string} 一覧用のラベル
     */
    function buildListLabel(text) {
        var listLabel = text.replace(/[\r\n\x03]+/g, " ");
        if (listLabel.length > LIST_LABEL_MAX_LENGTH) listLabel = listLabel.substring(0, LIST_LABEL_MAX_LENGTH) + "…";
        return listLabel;
    }

    /**
     * 編集欄の文字列を、書き戻す文字列にする
     * 編集欄は改行を \n（Windows は \r\n）で返すので、段落の改行 \r にそろえる。強制改行の代替文字は \x03 に戻す
     * @param {string} editText - 編集欄の文字列
     * @returns {string} 書き戻す文字列
     */
    function toFrameText(editText) {
        return editText.replace(/\r\n|\n/g, '\r').split(SOFT_BREAK).join('\x03');
    }

    /**
     * 強制改行（\x03）を編集欄用の代替文字にする
     * @param {string} frameText - テキストの内容
     * @returns {string} 編集欄に出す文字列
     */
    function toEditText(frameText) {
        return frameText.replace(/\x03/g, SOFT_BREAK);
    }

    /**
     * 2桁にゼロ埋めする
     * @param {number} num - 数値
     * @returns {string} 2桁の文字列
     */
    function zeroPad2(num) {
        return (num < 10 ? '0' : '') + num;
    }

    /**
     * 現在日時を YYYYMMDD-HHMMSS 形式で返す
     * @returns {string} 日時の文字列
     */
    function getDateTimeStamp() {
        var now = new Date();
        return now.getFullYear()
            + zeroPad2(now.getMonth() + 1)
            + zeroPad2(now.getDate())
            + '-'
            + zeroPad2(now.getHours())
            + zeroPad2(now.getMinutes())
            + zeroPad2(now.getSeconds());
    }

    /**
     * ファイル名に使えない文字を _ に置き換える
     * @param {string} fileName - ファイル名
     * @returns {string} 置き換えたファイル名
     */
    function sanitizeFileName(fileName) {
        return fileName.replace(/[\\\/\:\*\?\"\<\>\|]+/g, '_');
    }

    // =========================================
    // 収集の条件 / Collection filters
    // =========================================

    /**
     * 実質的に空のテキストか（改行・空白だけのものも空とみなす）
     * @param {TextFrame} textFrame - 調べるテキストフレーム
     * @returns {boolean} 空なら true
     */
    function isEmptyTextFrame(textFrame) {
        return textFrame.contents.replace(/[\s\x03]+/g, "").length === 0;
    }

    /**
     * レイヤー名が // ではじまるか
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {boolean} // ではじまるレイヤーにあれば true
     */
    function isCommentLayer(pageItem) {
        return pageItem.layer.name.indexOf("//") === 0;
    }

    /**
     * 非表示・ロックの設定に照らして収集してよいか（自身と親のグループ・レイヤーを見る）
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden, includeSymbols }
     * @returns {boolean} 収集してよければ true
     */
    function isCollectable(pageItem, collectOptions) {
        for (var node = pageItem; node && node.typename !== "Document"; node = node.parent) {
            if (!collectOptions.includeLocked && node.locked) return false;
            /* レイヤーは visible、オブジェクトは hidden で持つ / Layers use visible, page items use hidden */
            var isHidden = (node.typename === "Layer") ? !node.visible : node.hidden;
            if (!collectOptions.includeHidden && isHidden) return false;
        }
        return true;
    }

    /**
     * アートボードと一部でも重なっているか
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @param {number[]} artboardRect - アートボードの [左, 上, 右, 下]
     * @returns {boolean} 重なっていれば true
     */
    function isOnArtboard(pageItem, artboardRect) {
        var bounds = pageItem.geometricBounds;
        return (bounds[2] > artboardRect[0] && bounds[0] < artboardRect[2] &&
            bounds[1] > artboardRect[3] && bounds[3] < artboardRect[1]);
    }

    /**
     * 重なっている最初のアートボードの番号を返す
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {number} アートボードの番号（0 始まり）。どこにも重ならなければ -1
     */
    function getItemArtboardIndex(doc, pageItem) {
        for (var i = 0; i < doc.artboards.length; i++) {
            if (isOnArtboard(pageItem, doc.artboards[i].artboardRect)) {
                return i;
            }
        }
        return -1;
    }

    /**
     * アートボードの番号が範囲に入るか
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @param {number} artboardIndex - 調べる番号（どこにも無ければ -1）
     * @param {number} activeArtboardIndex - 現在のアートボードの番号
     * @returns {boolean} 範囲に入れば true
     */
    function isArtboardIndexInScope(scopeMode, artboardIndex, activeArtboardIndex) {
        if (scopeMode === 'current') return artboardIndex === activeArtboardIndex;
        if (scopeMode === 'allArtboards') return artboardIndex >= 0;
        return true;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。使わない関数も消さずに残してよい（互いに呼び合う）。
    //    識別子は SELECTION_ITEMS_TOLERANCE / normalizeSelectionItems / resolveTextRangeFrame /
    //    collectSelectionItems / getTextFrameKindKey / collectSelectionTextFrames / collectSelectionPathItems /
    //    isClipMaskItem / getClipMaskItem / hasClippedDescendant / readUsePreviewBoundsPreference /
    //    getClipAwareBounds / filterMeasurableChildren / getClipAwareUnionBounds / isNearlySameCoordinate / areBoundsNearlyEqual
    // 2. 選択は normalizeSelectionItems(doc.selection) で配列にする。文字カーソルの選択（TextRange）は
    //    配列ではなく1個で返り、しかも .length（文字数）を持つので、length だけで配列と見なさない
    // 3. テキストフレーム:
    //      var frames = collectSelectionTextFrames(doc.selection);                           // 全種類
    //      var frames = collectSelectionTextFrames(doc.selection, { kinds: ["point", "path"] });
    //    パス:
    //      var paths = collectSelectionPathItems(doc.selection);                             // 複合パスは中のパスへ
    //      var paths = collectSelectionPathItems(doc.selection, { compoundPaths: "whole", skipClipMasks: true });
    //    それ以外は collectSelectionItems(source, { accept: function (item) { … } }) で条件を書く
    // 4. 並びは選択と同じ前面→背面（グループの中も pageItems の順）。重なり順を使う処理はこの順を前提にしてよい
    // 5. doc.selection に代入し直す配列は skipLocked / skipHidden を true にする。
    //    ロック・非表示を選択に代入すると例外になり、中の子が選択に残る
    // 6. 境界は getClipAwareBounds(item, usePreviewBounds) / getClipAwareUnionBounds(items, usePreviewBounds)。
    //    usePreviewBounds を省くと環境設定の［プレビュー境界を使用］に従う。返り値は [左, 上, 右, 下] の新しい配列
    //    （書き換えても元のオブジェクトに影響しない）。測れないときは null
    // 7. 座標の一致・前後の判定は isNearlySameCoordinate / areBoundsNearlyEqual で許容値を挟む
    //    （吸着させた辺とガイドは 1e-12 ほどずれる）
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /* 座標を同じと見なす許容値（pt） / Tolerance for treating coordinates as equal, in points */
    var SELECTION_ITEMS_TOLERANCE = 0.001;

    /**
     * 選択やコレクションを、オブジェクトの配列にそろえる
     * TextRange・PathItem は length を持つので、typename で1個か集まりかを見分ける
     * @param {*} source - doc.selection、配列、DOM のコレクション、または単独のオブジェクト
     * @returns {Array} オブジェクトの配列（空なら []）
     */
    function normalizeSelectionItems(source) {
        var items = [];
        if (!source) return items;
        var typeName = "";
        try { typeName = source.typename || ""; } catch (e) { /* 読めない種類 / unreadable kind */ }
        /* 単数形の typename は1個（PageItems などのコレクションは s で終わる）
           A singular typename is one object (collections such as PageItems end in s) */
        if (typeName && !/s$/.test(typeName)) return [source];
        if (typeof source.length !== "number") return items;
        for (var i = 0; i < source.length; i++) items.push(source[i]);
        return items;
    }

    /**
     * 文字カーソルの選択（TextRange）を、それを含むテキストフレームに読み替える
     * @param {TextRange} textRange - 文字の範囲
     * @returns {TextFrame|null} テキストフレーム（たどれなければ null）
     */
    function resolveTextRangeFrame(textRange) {
        var current = textRange;
        /* parent をたどる（深さは念のため制限） / Walk up the parents, with a safety limit */
        for (var depth = 0; depth < 10 && current; depth++) {
            try {
                if (current.typename === "TextFrame") return current;
                current = current.parent;
            } catch (e) {
                break;
            }
        }
        /* ストーリーの先頭フレームで代用する / Fall back to the first frame of the story */
        try {
            var storyFrames = textRange.story.textFrames;
            if (storyFrames.length > 0) return storyFrames[0];
        } catch (e2) { /* ストーリーを持たない / no story */ }
        return null;
    }

    /**
     * 選択から条件に合うオブジェクトを集める（グループ・レイヤーを再帰でたどり、重複は除く）
     * 条件に合ったオブジェクトの中へは進まない
     * @param {*} source - doc.selection、配列、コレクション、または単独のオブジェクト
     * @param {Object} [options] - 収集の設定
     * @param {function(PageItem): boolean} [options.accept] - 集める条件（既定はグループ・レイヤー以外すべて）
     * @param {boolean} [options.enterGroups] - グループの中をたどる（既定 true）
     * @param {boolean} [options.enterClipGroups] - クリップグループの中をたどる（既定は enterGroups と同じ）
     * @param {boolean} [options.enterCompoundPaths] - 複合パスの中のパスをたどる（既定 false）
     * @param {boolean} [options.textRangeToFrame] - 文字の選択をテキストフレームに読み替える（既定 true）
     * @param {boolean} [options.skipLocked] - ロックされたものを中ごと外す（既定 false）
     * @param {boolean} [options.skipHidden] - 非表示のものを中ごと外す（既定 false）
     * @param {boolean} [options.skipClipMasks] - クリッピングマスクを外す（既定 false）
     * @param {boolean} [options.skipGuides] - ガイドを外す（既定 false）
     * @param {boolean} [options.unique] - 同じ参照を1回だけにする（既定 true。数千件で遅ければ false）
     * @returns {Array} 集めたオブジェクト（前面→背面の順）
     */
    function collectSelectionItems(source, options) {
        var opts = options || {};
        var enterGroups = (opts.enterGroups !== false);
        var enterClipGroups = (opts.enterClipGroups === undefined) ? enterGroups : (opts.enterClipGroups === true);
        var accept = opts.accept || function (item) {
            return item.typename !== "GroupItem" && item.typename !== "Layer";
        };
        var collected = [];

        /**
         * 集めた配列に加える（unique のときは同じ参照を足さない）
         * @param {PageItem} item - 加えるオブジェクト
         * @returns {void}
         */
        function pushItem(item) {
            if (opts.unique !== false) {
                for (var k = 0; k < collected.length; k++) {
                    if (collected[k] === item) return;
                }
            }
            collected.push(item);
        }

        /**
         * 設定に従って外すオブジェクトか判定する
         * @param {PageItem} item - 判定するオブジェクト
         * @returns {boolean} 外すなら true
         */
        function isSkipped(item) {
            try {
                if (item.typename === "Layer") {
                    if (opts.skipLocked && item.locked) return true;
                    if (opts.skipHidden && !item.visible) return true;
                    return false;
                }
                if (opts.skipLocked && item.locked) return true;
                if (opts.skipHidden && item.hidden) return true;
                if (opts.skipGuides && item.guides === true) return true;
                if (opts.skipClipMasks && isClipMaskItem(item)) return true;
            } catch (e) {
                /* 読めないプロパティは「外さない」に倒す / Unreadable properties do not exclude */
            }
            return false;
        }

        /**
         * 1件をたどって集める
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {void}
         */
        function visit(item) {
            if (!item) return;
            var typeName = "";
            try { typeName = item.typename; } catch (e) { return; }

            if (typeName === "TextRange" || typeName === "InsertionPoint") {
                if (opts.textRangeToFrame === false) {
                    if (accept(item)) pushItem(item);
                    return;
                }
                visit(resolveTextRangeFrame(item));
                return;
            }
            if (isSkipped(item)) return;
            if (accept(item)) {
                pushItem(item);
                return;
            }

            var children = null;
            if (typeName === "GroupItem") {
                var isClipped = false;
                try { isClipped = (item.clipped === true); } catch (e2) { }
                if (isClipped ? enterClipGroups : enterGroups) children = item.pageItems;
            } else if (typeName === "CompoundPathItem") {
                if (opts.enterCompoundPaths) children = item.pathItems;
            } else if (typeName === "Layer") {
                /* 重なり順はサブレイヤーとページアイテムで別々なので、ページアイテム→サブレイヤーの順にする
                   Page items and sublayers stack separately; visit page items first, then sublayers */
                walk(item.pageItems);
                walk(item.layers);
                return;
            }
            if (children) walk(children);
        }

        /**
         * 集まりの各要素をたどる
         * @param {*} list - 配列またはコレクション
         * @returns {void}
         */
        function walk(list) {
            var listItems = normalizeSelectionItems(list);
            for (var i = 0; i < listItems.length; i++) visit(listItems[i]);
        }

        walk(source);
        return collected;
    }

    /**
     * テキストフレームの種類を "point" / "area" / "path" で返す
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {string} 種類のキー（判定できなければ ""）
     */
    function getTextFrameKindKey(textFrame) {
        try {
            if (textFrame.kind === TextType.POINTTEXT) return "point";
            if (textFrame.kind === TextType.AREATEXT) return "area";
            if (textFrame.kind === TextType.PATHTEXT) return "path";
        } catch (e) { /* kind を読めない / kind is unreadable */ }
        return "";
    }

    /**
     * 選択からテキストフレームを集める（グループの中・文字カーソルの選択を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string[]} [options.kinds] - 集める種類（"point" / "area" / "path"。既定はすべて）
     * @returns {TextFrame[]} テキストフレーム（前面→背面の順）
     */
    function collectSelectionTextFrames(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var kindFilter = null;
        if (opts.kinds && opts.kinds.length) {
            kindFilter = {};
            for (var i = 0; i < opts.kinds.length; i++) kindFilter[opts.kinds[i]] = true;
        }
        opts.accept = function (item) {
            if (item.typename !== "TextFrame") return false;
            return !kindFilter || kindFilter[getTextFrameKindKey(item)] === true;
        };
        /* 種類で外したテキストは中をたどらない（accept が false でも子は無い） / Text frames have no children to walk */
        return collectSelectionItems(source, opts);
    }

    /**
     * 選択からパスを集める（グループの中を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string} [options.compoundPaths] - 複合パスの扱い。"children"（中のパス、既定）/ "whole"（複合パスごと）/ "skip"（外す）
     * @returns {Array} PathItem（"whole" のときは CompoundPathItem も）の配列
     */
    function collectSelectionPathItems(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var compoundMode = opts.compoundPaths || "children";
        opts.enterCompoundPaths = (compoundMode === "children");
        opts.accept = function (item) {
            if (item.typename === "PathItem") return true;
            return compoundMode === "whole" && item.typename === "CompoundPathItem";
        };
        return collectSelectionItems(source, opts);
    }

    /**
     * クリッピングマスク（クリップグループの型）か判定する
     * パスは clipping、複合パスは中の先頭パスの clipping、テキストは clipping が無いので「クリップグループの先頭」で見る
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} マスクなら true
     */
    function isClipMaskItem(item) {
        try {
            if (item.typename === "PathItem") return item.clipping === true;
            if (item.typename === "CompoundPathItem") {
                return item.pathItems.length > 0 && item.pathItems[0].clipping === true;
            }
            if (item.typename === "TextFrame") {
                var parentGroup = item.parent;
                return parentGroup.typename === "GroupItem" && parentGroup.clipped === true &&
                    parentGroup.pageItems.length > 0 && parentGroup.pageItems[0] === item;
            }
        } catch (e) { /* 読めない種類はマスクではない / unreadable kinds are not masks */ }
        return false;
    }

    /**
     * クリップグループの型（マスク）を返す
     * フラグで探し、見つからなければ先頭（pageItems[0]）を返す（型は常に最前面。テキストの型はフラグを持たない）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PageItem|null} マスク（クリップグループでなければ null）
     */
    function getClipMaskItem(groupItem) {
        try {
            if (!groupItem || groupItem.typename !== "GroupItem" || groupItem.clipped !== true) return null;
            var groupChildren = groupItem.pageItems;
            if (groupChildren.length === 0) return null;
            for (var i = 0; i < groupChildren.length; i++) {
                var childType = groupChildren[i].typename;
                if ((childType === "PathItem" || childType === "CompoundPathItem") && isClipMaskItem(groupChildren[i])) {
                    return groupChildren[i];
                }
            }
            return groupChildren[0];
        } catch (e) {
            return null;
        }
    }

    /**
     * グループの中（入れ子を含む）にクリップグループがあるか判定する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {boolean} あれば true
     */
    function hasClippedDescendant(groupItem) {
        try {
            var groupChildren = groupItem.pageItems;
            for (var i = 0; i < groupChildren.length; i++) {
                if (groupChildren[i].typename !== "GroupItem") continue;
                if (groupChildren[i].clipped === true || hasClippedDescendant(groupChildren[i])) return true;
            }
        } catch (e) { /* 中を読めない / cannot read the children */ }
        return false;
    }

    /**
     * 環境設定の［プレビュー境界を使用］を読む
     * @returns {boolean} オンなら true（読めなければ false）
     */
    function readUsePreviewBoundsPreference() {
        try {
            return app.preferences.getBooleanPreference("includeStrokeInBounds");
        } catch (e) {
            return false;
        }
    }

    /**
     * 見た目どおりの境界を返す。クリップグループはマスクの境界、
     * 中にクリップグループを含むグループは子の境界を合わせたもの（隠れた部分を含めない）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下] の新しい配列（測れなければ null）
     */
    function getClipAwareBounds(item, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        try {
            var measuredItem = item;
            if (item.typename === "GroupItem") {
                var maskItem = getClipMaskItem(item);
                if (maskItem) {
                    measuredItem = maskItem;
                } else if (hasClippedDescendant(item)) {
                    /* グループ自体の効果（影など）の広がりは含まれなくなる
                       This leaves out the reach of effects applied to the group itself (drop shadows etc.) */
                    var childBounds = getClipAwareUnionBounds(filterMeasurableChildren(item.pageItems), usePreview);
                    if (childBounds) return childBounds;
                }
            }
            var bounds = usePreview ? measuredItem.visibleBounds : measuredItem.geometricBounds;
            return [bounds[0], bounds[1], bounds[2], bounds[3]];
        } catch (e) {
            return null;
        }
    }

    /**
     * 境界の計算に入れる子だけを残す（非表示とガイドを外す）
     * @param {*} childList - 子のコレクション
     * @returns {Array} 残した子
     */
    function filterMeasurableChildren(childList) {
        var childItems = normalizeSelectionItems(childList);
        var measurable = [];
        for (var i = 0; i < childItems.length; i++) {
            try {
                if (childItems[i].hidden === true || childItems[i].guides === true) continue;
            } catch (e) { /* 読めなければ残す / keep when unreadable */ }
            measurable.push(childItems[i]);
        }
        return measurable;
    }

    /**
     * 複数のオブジェクトを囲む外接範囲を返す（クリップグループはマスクで測る）
     * @param {*} items - オブジェクトの配列・コレクション・選択
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下]（測れるものが無ければ null）
     */
    function getClipAwareUnionBounds(items, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        var itemList = normalizeSelectionItems(items);
        var unionBounds = null;
        for (var i = 0; i < itemList.length; i++) {
            var itemBounds = getClipAwareBounds(itemList[i], usePreview);
            if (!itemBounds) continue;
            if (!unionBounds) {
                unionBounds = itemBounds;
                continue;
            }
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 2つの座標を許容値つきで比べる
     * @param {number} valueA - 座標A（pt）
     * @param {number} valueB - 座標B（pt）
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 差が許容値以下なら true
     */
    function isNearlySameCoordinate(valueA, valueB, tolerance) {
        var limit = (typeof tolerance === "number") ? tolerance : SELECTION_ITEMS_TOLERANCE;
        return Math.abs(valueA - valueB) <= limit;
    }

    /**
     * 2つの境界を許容値つきで比べる
     * @param {number[]} boundsA - [左, 上, 右, 下]
     * @param {number[]} boundsB - [左, 上, 右, 下]
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 4辺とも許容値以内なら true
     */
    function areBoundsNearlyEqual(boundsA, boundsB, tolerance) {
        if (!boundsA || !boundsB) return false;
        for (var i = 0; i < 4; i++) {
            if (!isNearlySameCoordinate(boundsA[i], boundsB[i], tolerance)) return false;
        }
        return true;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // テキストの収集 / Collecting text frames
    // =========================================

    /**
     * レイヤーまたはグループの直下からテキストを集め、グループは中へたどる
     * @param {Object} container - レイヤーまたはグループ
     * @param {TextFrame[]} collectedFrames - 集めたテキストフレームを入れる配列
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden, includeSymbols }
     * @param {Function} acceptFrame - テキストフレームを受け入れるか判定する関数
     * @returns {void}
     */
    function collectTextFromContainer(container, collectedFrames, collectOptions, acceptFrame) {
        var pageItems = container.pageItems;
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (pageItem.parent !== container) continue;
            if (!collectOptions.includeComment && isCommentLayer(pageItem)) continue;
            if (!isCollectable(pageItem, collectOptions)) continue;

            if (pageItem.typename === "TextFrame") {
                if (!isEmptyTextFrame(pageItem) && acceptFrame(pageItem)) {
                    collectedFrames.push(pageItem);
                }
            } else if (pageItem.typename === "GroupItem") {
                collectTextFromContainer(pageItem, collectedFrames, collectOptions, acceptFrame);
            }
        }
    }

    /**
     * レイヤーとそのサブレイヤーからテキストを集める
     * @param {Layer} layer - 対象のレイヤー
     * @param {TextFrame[]} collectedFrames - 集めたテキストフレームを入れる配列
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden, includeSymbols }
     * @param {Function} acceptFrame - テキストフレームを受け入れるか判定する関数
     * @returns {void}
     */
    function collectTextFromLayer(layer, collectedFrames, collectOptions, acceptFrame) {
        collectTextFromContainer(layer, collectedFrames, collectOptions, acceptFrame);
        for (var i = 0; i < layer.layers.length; i++) {
            collectTextFromLayer(layer.layers[i], collectedFrames, collectOptions, acceptFrame);
        }
    }

    /**
     * 範囲に応じてテキストフレームを集める
     * @param {Document} doc - 対象のドキュメント
     * @param {string} scopeMode - "all"（ドキュメント全体）/ "allArtboards"（いずれかのアートボード）/ "current"（現在のアートボード）
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden, includeSymbols }
     * @returns {TextFrame[]} テキストフレームの配列
     */
    function collectFramesByScope(doc, scopeMode, collectOptions) {
        var collectedFrames = [];
        var acceptFrame;
        if (scopeMode === "all") {
            acceptFrame = function () { return true; };
        } else if (scopeMode === "allArtboards") {
            acceptFrame = function (textFrame) { return getItemArtboardIndex(doc, textFrame) >= 0; };
        } else {
            var activeArtboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            acceptFrame = function (textFrame) { return isOnArtboard(textFrame, activeArtboardRect); };
        }
        for (var i = 0; i < doc.layers.length; i++) {
            collectTextFromLayer(doc.layers[i], collectedFrames, collectOptions, acceptFrame);
        }
        return collectedFrames;
    }

    /**
     * 名前一覧タブ用に、ドキュメント全体のテキストを条件なしで集める
     * @param {Document} doc - 対象のドキュメント
     * @returns {TextFrame[]} テキストフレームの配列
     */
    function collectAllTextFrames(doc) {
        return collectFramesByScope(doc, 'all', {
            includeComment: true,
            includeLocked: true,
            includeHidden: true
        });
    }

    /**
     * 同じ内容のテキストフレームをまとめる
     * @param {TextFrame[]} textFrames - テキストフレームの配列
     * @returns {{uniqueFrames: TextFrame[], duplicateMap: Array}} 内容ごとの代表と、代表ごとの同じ内容の全フレーム
     */
    function groupDuplicateFrames(textFrames) {
        var uniqueFrames = [];
        var duplicateMap = [];
        var contentIndexMap = {}; /* 内容 → uniqueFrames の番号 / contents -> index in uniqueFrames */
        for (var i = 0; i < textFrames.length; i++) {
            var frameContents = textFrames[i].contents;
            if (contentIndexMap.hasOwnProperty(frameContents)) {
                duplicateMap[contentIndexMap[frameContents]].push(textFrames[i]);
            } else {
                contentIndexMap[frameContents] = uniqueFrames.length;
                uniqueFrames.push(textFrames[i]);
                duplicateMap.push([textFrames[i]]);
            }
        }
        return { uniqueFrames: uniqueFrames, duplicateMap: duplicateMap };
    }

    // =========================================
    // レイヤー名・アートボード名 / Layer and artboard names
    // =========================================

    /**
     * レイヤーの名前を重複なく集める
     * @param {Layers} layers - 対象のレイヤーのコレクション
     * @param {boolean} includeSublayers - サブレイヤーもたどるなら true
     * @param {string[]} layerNames - 集めた名前を入れる配列
     * @param {Object} seenNames - 既出の名前の表
     * @returns {void}
     */
    function collectLayerNames(layers, includeSublayers, layerNames, seenNames) {
        for (var i = 0; i < layers.length; i++) {
            var layerName = layers[i].name;
            if (layerName && !seenNames.hasOwnProperty(layerName)) {
                seenNames[layerName] = true;
                layerNames.push(layerName);
            }
            if (includeSublayers) {
                collectLayerNames(layers[i].layers, true, layerNames, seenNames);
            }
        }
    }

    /**
     * ドキュメントのレイヤー名を重複なく集める
     * @param {Document} doc - 対象のドキュメント
     * @param {boolean} includeSublayers - サブレイヤーも含めるなら true
     * @returns {string[]} レイヤー名の配列
     */
    function collectDocumentLayerNames(doc, includeSublayers) {
        var layerNames = [];
        collectLayerNames(doc.layers, includeSublayers, layerNames, {});
        return layerNames;
    }

    /**
     * アートボード番号の既定名（「アートボード1」「Artboard 1」）を返す
     * @param {number} index - アートボードの番号（0 始まり）
     * @returns {string} 既定名
     */
    function getArtboardNumberName(index) {
        return getLabel('fallbackName.artboardNumber').replace('{number}', index + 1);
    }

    /**
     * アートボード名を返す（空なら既定名）
     * @param {Document} doc - 対象のドキュメント
     * @param {number} index - アートボードの番号（0 始まり）
     * @returns {string} アートボード名
     */
    function getArtboardName(doc, index) {
        return doc.artboards[index].name || getArtboardNumberName(index);
    }

    /**
     * ［アートボード名］タブの「番号つき」表示（「1: 名前」）を返す
     * @param {Document} doc - 対象のドキュメント
     * @param {number} index - アートボードの番号（0 始まり）
     * @returns {string} 表示名
     */
    function getNumberedArtboardName(doc, index) {
        return (index + 1) + ': ' + getArtboardName(doc, index);
    }

    /**
     * 書き出し用の見出し（「アートボード1: 名前」）を返す
     * @param {Document} doc - 対象のドキュメント
     * @param {number} index - アートボードの番号（0 始まり）
     * @returns {string} 見出し
     */
    function getArtboardExportHeading(doc, index) {
        return getArtboardNumberName(index) + ': ' + getArtboardName(doc, index);
    }

    // =========================================
    // フォント / Fonts
    // =========================================

    /**
     * フォントの組（PostScript 名・ファミリー名・スタイル名）の比較用キーを返す
     * @param {Object} fontTriple - { psName, familyName, styleName }
     * @returns {string} タブ区切りのキー
     */
    function getFontTripleKey(fontTriple) {
        return fontTriple.psName + '\t' + fontTriple.familyName + '\t' + fontTriple.styleName;
    }

    /**
     * 文字のフォントの組を返す（読めなければ「不明」）
     * @param {TextRange} textCharacter - 対象の文字
     * @returns {Object} { psName, familyName, styleName }
     */
    function getCharacterFontTriple(textCharacter) {
        var unknownFont = getLabel('fallbackName.unknownFont');
        /* 環境にないフォントは読み取りで例外になることがある / Missing fonts may throw when read */
        try {
            var textFont = textCharacter.characterAttributes.textFont;
            return {
                psName: textFont.name || unknownFont,
                familyName: textFont.family || unknownFont,
                styleName: textFont.style || ''
            };
        } catch (e) {
            return { psName: unknownFont, familyName: unknownFont, styleName: '' };
        }
    }

    /**
     * 複数のテキストで使われているフォントの組を、出てきた順に重複なく集める
     * @param {TextFrame[]} textFrames - テキストフレームの配列
     * @returns {Object[]} { psName, familyName, styleName } の配列
     */
    function collectUniqueFontTriples(textFrames) {
        var seenKeys = {};
        var uniqueTriples = [];
        for (var i = 0; i < textFrames.length; i++) {
            var frameCharacters = textFrames[i].characters;
            for (var j = 0; j < frameCharacters.length; j++) {
                var fontTriple = getCharacterFontTriple(frameCharacters[j]);
                var tripleKey = getFontTripleKey(fontTriple);
                if (!seenKeys.hasOwnProperty(tripleKey)) {
                    seenKeys[tripleKey] = true;
                    uniqueTriples.push(fontTriple);
                }
            }
        }
        return uniqueTriples;
    }

    /**
     * フォントの組の配列を、出てきた順に重複なく1つにまとめる
     * @param {Object[][]} tripleLists - { psName, familyName, styleName } の配列の配列
     * @returns {Object[]} まとめた配列
     */
    function mergeFontTriples(tripleLists) {
        var seenKeys = {};
        var mergedTriples = [];
        for (var i = 0; i < tripleLists.length; i++) {
            for (var j = 0; j < tripleLists[i].length; j++) {
                var tripleKey = getFontTripleKey(tripleLists[i][j]);
                if (!seenKeys.hasOwnProperty(tripleKey)) {
                    seenKeys[tripleKey] = true;
                    mergedTriples.push(tripleLists[i][j]);
                }
            }
        }
        return mergedTriples;
    }

    /**
     * フォントの組の配列に、指定の PostScript 名があるか
     * @param {Object[]} fontTriples - { psName, familyName, styleName } の配列
     * @param {string} targetPSName - フォントの PostScript 名
     * @returns {boolean} あれば true
     */
    function containsFont(fontTriples, targetPSName) {
        for (var i = 0; i < fontTriples.length; i++) {
            if (fontTriples[i].psName === targetPSName) return true;
        }
        return false;
    }

    /**
     * テキストが指定のフォントを使っているか
     * @param {TextFrame} textFrame - 調べるテキストフレーム
     * @param {string} targetPSName - フォントの PostScript 名
     * @returns {boolean} 使っていれば true
     */
    function usesFont(textFrame, targetPSName) {
        var frameCharacters = textFrame.characters;
        for (var i = 0; i < frameCharacters.length; i++) {
            if (getCharacterFontTriple(frameCharacters[i]).psName === targetPSName) return true;
        }
        return false;
    }

    /**
     * 書き出し用に、フォントの行（「PostScript名  ファミリー名  スタイル名」）を並べ替えて返す
     * @param {Object[]} fontTriples - { psName, familyName, styleName } の配列
     * @returns {string[]} フォントの行
     */
    function buildFontExportLines(fontTriples) {
        var fontLines = [];
        for (var i = 0; i < fontTriples.length; i++) {
            fontLines.push(fontTriples[i].psName + '  ' + fontTriples[i].familyName + (fontTriples[i].styleName ? ('  ' + fontTriples[i].styleName) : ''));
        }
        fontLines.sort();
        return fontLines;
    }

    /**
     * 指定のフォントを使っているテキストと、そのフォントを含むシンボルのインスタンスを選択する（ドキュメント全体から探す）
     * @param {Document} doc - 対象のドキュメント
     * @param {string} targetPSName - フォントの PostScript 名
     * @param {Object[]} symbolEntries - ドキュメント全体のシンボル内テキスト（readSymbolTextEntries() の結果）
     * @returns {void}
     */
    function selectItemsByFont(doc, targetPSName, symbolEntries) {
        var textFrames = collectAllTextFrames(doc);
        var matchedItems = [];
        var matchedSymbols = [];
        var i;

        for (i = 0; i < textFrames.length; i++) {
            if (usesFont(textFrames[i], targetPSName)) matchedItems.push(textFrames[i]);
        }
        for (i = 0; i < symbolEntries.length; i++) {
            if (containsFont(symbolEntries[i].fontTriples, targetPSName) && indexOfItem(matchedSymbols, symbolEntries[i].symbol) === -1) {
                matchedSymbols.push(symbolEntries[i].symbol);
            }
        }
        if (matchedSymbols.length > 0) {
            for (i = 0; i < doc.symbolItems.length; i++) {
                if (indexOfItem(matchedSymbols, doc.symbolItems[i].symbol) !== -1) matchedItems.push(doc.symbolItems[i]);
            }
        }

        doc.selection = null;
        if (matchedItems.length === 0) {
            alert(getLabel('alert.noMatchingFontText'));
            return;
        }

        for (i = 0; i < matchedItems.length; i++) {
            /* ロック・非表示のものは選択できない / Locked or hidden items cannot be selected */
            try {
                matchedItems[i].selected = true;
            } catch (e) { }
        }
        app.redraw();
    }

    // =========================================
    // 書式を保持した書き換え / Rewriting while keeping formatting
    // =========================================

    /**
     * 自身と親の非表示・ロックを一時的に解除し、元に戻すための記録を追加する
     * @param {PageItem} pageItem - 対象のアイテム
     * @param {Object[]} stateRecords - { target, propertyName, originalValue } の追加先
     * @returns {void}
     */
    function unlockAndRevealAncestors(pageItem, stateRecords) {
        for (var node = pageItem; node && node.typename !== "Document"; node = node.parent) {
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
     * 自身と親のグループ・レイヤーを一時的に編集できる状態にして処理し、終わったら元に戻す
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {Function} editAction - 編集できる状態で行う処理
     * @returns {void}
     */
    function withTemporarilyEditableItem(pageItem, editAction) {
        var stateRecords = [];
        try {
            unlockAndRevealAncestors(pageItem, stateRecords);
            editAction();
        } finally {
            restoreItemStates(stateRecords);
        }
    }

    /**
     * 文字コードが UTF-16 の下位サロゲートか（絵文字などを途中で切らないための判定）
     * @param {number} charCode - 文字コード
     * @returns {boolean} 下位サロゲートなら true
     */
    function isLowSurrogate(charCode) {
        return charCode >= 0xDC00 && charCode <= 0xDFFF;
    }

    /**
     * 新旧の文字列を先頭と末尾から突き合わせ、変わった範囲を返す
     * @param {string} oldText - 元の内容
     * @param {string} newText - 新しい内容
     * @returns {{start: number, oldEnd: number, newMiddle: string}} 元の内容で書き換える範囲 [start, oldEnd) と、そこに入る文字列
     */
    function findChangedRange(oldText, newText) {
        var commonLimit = Math.min(oldText.length, newText.length);
        var prefixLength = 0;
        while (prefixLength < commonLimit && oldText.charAt(prefixLength) === newText.charAt(prefixLength)) {
            prefixLength++;
        }
        /* サロゲートペアの間で切らない / Do not split a surrogate pair */
        if (prefixLength > 0 && isLowSurrogate(oldText.charCodeAt(prefixLength))) prefixLength--;

        var suffixLength = 0;
        while (suffixLength < commonLimit - prefixLength &&
            oldText.charAt(oldText.length - 1 - suffixLength) === newText.charAt(newText.length - 1 - suffixLength)) {
            suffixLength++;
        }
        if (suffixLength > 0 && isLowSurrogate(oldText.charCodeAt(oldText.length - suffixLength))) suffixLength--;

        return {
            start: prefixLength,
            oldEnd: oldText.length - suffixLength,
            newMiddle: newText.substring(prefixLength, newText.length - suffixLength)
        };
    }

    /**
     * 変わった文字だけを書き換える（contents 全体を書き戻さないので、文字・段落の書式が残る）
     * 挿入した文字は、行頭なら後ろの文字、それ以外は前の文字の書式を引き継ぐ
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} newText - 新しい内容
     * @returns {void}
     */
    function replaceChangedCharacters(textFrame, newText) {
        var oldText = textFrame.contents;
        if (oldText === newText) return;
        if (oldText.length === 0) {
            textFrame.contents = newText;
            return;
        }

        var changedRange = findChangedRange(oldText, newText);
        var frameCharacters = textFrame.characters;

        if (changedRange.start < changedRange.oldEnd) {
            /* 右から消し、残した1文字目を新しい文字列に差し替える / Remove from the right, then replace the first changed character */
            for (var i = changedRange.oldEnd - 1; i > changedRange.start; i--) {
                frameCharacters[i].remove();
            }
            if (changedRange.newMiddle) {
                frameCharacters[changedRange.start].contents = changedRange.newMiddle;
            } else {
                frameCharacters[changedRange.start].remove();
            }
            return;
        }

        /* 挿入だけのとき / Insertion only */
        var previousChar = (changedRange.start > 0) ? oldText.charAt(changedRange.start - 1) : '';
        var isLineStart = (previousChar === '' || /[\r\n\x03]/.test(previousChar));
        if (isLineStart && changedRange.start < oldText.length) {
            var nextCharacter = frameCharacters[changedRange.start];
            nextCharacter.contents = changedRange.newMiddle + nextCharacter.contents;
        } else {
            var previousCharacter = frameCharacters[changedRange.start - 1];
            previousCharacter.contents = previousCharacter.contents + changedRange.newMiddle;
        }
    }

    /**
     * テキストの内容を置き換える。書式を保持するときは変わった文字だけを書き換える
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} newText - 新しい内容
     * @param {boolean} keepFormat - 文字・段落の書式を保持するなら true
     * @returns {void}
     */
    function replaceTextContents(textFrame, newText, keepFormat) {
        withTemporarilyEditableItem(textFrame, function () {
            if (keepFormat) {
                replaceChangedCharacters(textFrame, newText);
            } else {
                textFrame.contents = newText;
            }
        });
    }

    // =========================================
    // ソート / Sorting
    // =========================================

    /**
     * 上から下へ、ほぼ同じ高さなら左から右へ並べ替える
     * @param {TextFrame[]} textFrames - 並べ替える配列（その場で並べ替える）
     * @returns {void}
     */
    function sortByPosition(textFrames) {
        textFrames.sort(function (firstFrame, secondFrame) {
            if (Math.abs(secondFrame.top - firstFrame.top) <= POSITION_SORT_TOLERANCE) {
                return firstFrame.left - secondFrame.left;
            }
            return secondFrame.top - firstFrame.top;
        });
    }

    /**
     * 内容の順（大文字小文字を区別しない）に並べ替える
     * @param {TextFrame[]} textFrames - 並べ替える配列（その場で並べ替える）
     * @returns {void}
     */
    function sortByContent(textFrames) {
        textFrames.sort(function (firstFrame, secondFrame) {
            var firstText = firstFrame.contents.toLowerCase();
            var secondText = secondFrame.contents.toLowerCase();
            if (firstText < secondText) return -1;
            if (firstText > secondText) return 1;
            return 0;
        });
    }

    // =========================================
    // シンボル内のテキスト / Text in symbols
    // =========================================

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
     * 範囲に入るシンボルインスタンスを、シンボルとアートボードの組ごとに1つずつ集める
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden, includeSymbols }
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @returns {Object[]} { symbol, artboardIndex } の配列
     */
    function collectScopedSymbolPlacements(doc, collectOptions, scopeMode) {
        var symbolItems = doc.symbolItems;
        var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();
        var placements = [];
        var seenSymbols = [];     /* 出てきたシンボル / Symbols seen so far */
        var seenArtboards = [];   /* シンボルごとに見たアートボード番号 / Artboard indexes seen per symbol */

        for (var i = 0; i < symbolItems.length; i++) {
            var symbolItem = symbolItems[i];
            if (!collectOptions.includeComment && isCommentLayer(symbolItem)) continue;
            if (!isCollectable(symbolItem, collectOptions)) continue;

            var artboardIndex = getItemArtboardIndex(doc, symbolItem);
            if (!isArtboardIndexInScope(scopeMode, artboardIndex, activeArtboardIndex)) continue;

            var symbolIndex = indexOfItem(seenSymbols, symbolItem.symbol);
            if (symbolIndex === -1) {
                symbolIndex = seenSymbols.length;
                seenSymbols.push(symbolItem.symbol);
                seenArtboards.push({});
            }
            if (seenArtboards[symbolIndex].hasOwnProperty(artboardIndex)) continue;
            seenArtboards[symbolIndex][artboardIndex] = true;
            placements.push({ symbol: symbolItem.symbol, artboardIndex: artboardIndex });
        }
        return placements;
    }

    /**
     * 配置の一覧から、シンボルを出てきた順に重複なく取り出す
     * @param {Object[]} placements - collectScopedSymbolPlacements() の結果
     * @returns {Symbol[]} シンボルの配列
     */
    function getUniqueSymbols(placements) {
        var uniqueSymbols = [];
        for (var i = 0; i < placements.length; i++) {
            if (indexOfItem(uniqueSymbols, placements[i].symbol) === -1) uniqueSymbols.push(placements[i].symbol);
        }
        return uniqueSymbols;
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
     * 選択しているテキストフレームとシンボルインスタンスを集める（グループ内も含む。文字を選択中ならそのストーリーのフレーム）
     * @param {Document} doc - 対象のドキュメント
     * @returns {{frames: TextFrame[], symbolItems: SymbolItem[]}} 選択中のテキストフレームとシンボルインスタンス
     */
    function collectSelectedTextItems(doc) {
        var selectedItems = { frames: [], symbolItems: [] };
        var selection = doc.selection;
        if (!selection) return selectedItems;
        if (selection.typename === "TextRange") {
            var storyFrames = selection.story.textFrames;
            for (var j = 0; j < storyFrames.length; j++) {
                selectedItems.frames.push(storyFrames[j]);
            }
            return selectedItems;
        }
        var items = toItemArray(selection);
        selectedItems.frames = collectSelectionTextFrames(items, { unique: false });
        selectedItems.symbolItems = collectSelectionItems(items, {
            accept: function (item) { return item.typename === "SymbolItem"; },
            unique: false
        });
        return selectedItems;
    }

    /**
     * 作業レイヤーを作って処理を実行し、終わったら作業レイヤーを消して選択とアクティブレイヤーを戻す
     * @param {Document} doc - 対象のドキュメント
     * @param {Function} work - (workLayer) を受け取る処理
     * @returns {void}
     */
    function withSymbolWorkLayer(doc, work) {
        var savedSelection = toItemArray(doc.selection);
        var savedActiveLayer = doc.activeLayer;
        var workLayer = doc.layers.add();
        workLayer.name = WORK_LAYER_NAME;
        try {
            work(workLayer);
        } finally {
            workLayer.remove();
            doc.activeLayer = savedActiveLayer;
            doc.selection = null;
            for (var i = 0; i < savedSelection.length; i++) {
                /* ロック・非表示になったものや文字の選択は戻せない / Locked, hidden or text selections cannot be restored */
                try {
                    savedSelection[i].selected = true;
                } catch (e) { }
            }
        }
    }

    /**
     * シンボルの中身を作業レイヤーに展開し、1つのグループにまとめて返す
     * @param {Document} doc - 対象のドキュメント
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
     * 展開したシンボルのテキストフレームを返す（並びは書き換え時の番号と一致する）
     * @param {GroupItem} contentGroup - expandSymbolToGroup() の結果
     * @returns {TextFrame[]} テキストフレームの配列
     */
    function getSymbolTextFrames(contentGroup) {
        return collectSelectionTextFrames(contentGroup, { unique: false });
    }

    /**
     * シンボル内の空でないテキストを集める（ドキュメントは変えない）
     * @param {Document} doc - 対象のドキュメント
     * @param {Symbol[]} targetSymbols - 対象のシンボル
     * @returns {Object[]} { symbol, symbolName, frameIndex, contents, fontTriples } の配列
     */
    function readSymbolTextEntries(doc, targetSymbols) {
        var symbolEntries = [];
        if (targetSymbols.length === 0) return symbolEntries;
        withSymbolWorkLayer(doc, function (workLayer) {
            for (var i = 0; i < targetSymbols.length; i++) {
                var contentGroup = expandSymbolToGroup(doc, targetSymbols[i], workLayer);
                var symbolFrames = getSymbolTextFrames(contentGroup);
                for (var j = 0; j < symbolFrames.length; j++) {
                    if (isEmptyTextFrame(symbolFrames[j])) continue;
                    symbolEntries.push({
                        symbol: targetSymbols[i],
                        symbolName: targetSymbols[i].name,
                        frameIndex: j,
                        contents: symbolFrames[j].contents,
                        fontTriples: collectUniqueFontTriples([symbolFrames[j]])
                    });
                }
                contentGroup.remove();
            }
        });
        return symbolEntries;
    }

    /**
     * シンボル内テキスト一覧用に、範囲に入るシンボルのテキストを集める
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden, includeSymbols }
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @returns {Object[]} { symbol, symbolName, frameIndex, contents, fontTriples } の配列
     */
    function collectSymbolTextEntries(doc, collectOptions, scopeMode) {
        var placements = collectScopedSymbolPlacements(doc, collectOptions, scopeMode);
        return readSymbolTextEntries(doc, getUniqueSymbols(placements));
    }

    /**
     * 書き出し用に、シンボル内テキストを「内容〈シンボル：名前〉」の形でアートボード番号つきで集める
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden, includeSymbols }
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @returns {Object[]} { artboardIndex, text, fontTriples } の配列
     */
    function collectSymbolTextsByArtboard(doc, collectOptions, scopeMode) {
        var placements = collectScopedSymbolPlacements(doc, collectOptions, scopeMode);
        var symbolEntries = readSymbolTextEntries(doc, getUniqueSymbols(placements));
        var exportEntries = [];

        for (var i = 0; i < placements.length; i++) {
            for (var j = 0; j < symbolEntries.length; j++) {
                if (symbolEntries[j].symbol !== placements[i].symbol) continue;
                exportEntries.push({
                    artboardIndex: placements[i].artboardIndex,
                    text: symbolEntries[j].contents + getLabel('format.symbolSuffix').split('{symbolName}').join(symbolEntries[j].symbolName),
                    fontTriples: symbolEntries[j].fontTriples
                });
            }
        }
        return exportEntries;
    }

    /**
     * 新しいシンボルを作り、元のシンボルのインスタンスをすべて差し替えて、元のシンボルと入れ替える
     * @param {Document} doc - 対象のドキュメント
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
        } catch (e) { }
        return true;
    }

    /**
     * シンボル内のテキストを書き換え、シンボルの定義を差し替える
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} symbolEntry - readSymbolTextEntries() の要素
     * @param {string} newText - 新しい内容
     * @param {boolean} keepFormat - 文字・段落の書式を保持するなら true
     * @returns {boolean} 差し替えられたら true
     */
    function replaceSymbolText(doc, symbolEntry, newText, keepFormat) {
        var isReplaced = false;
        withSymbolWorkLayer(doc, function (workLayer) {
            var contentGroup = expandSymbolToGroup(doc, symbolEntry.symbol, workLayer);
            var targetFrame = getSymbolTextFrames(contentGroup)[symbolEntry.frameIndex];
            if (targetFrame) {
                replaceTextContents(targetFrame, newText, keepFormat);
                isReplaced = replaceSymbolDefinition(doc, symbolEntry.symbol, contentGroup);
            }
            contentGroup.remove();
        });
        return isReplaced;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // テキスト書き出し / Text export
    // =========================================

    /**
     * 範囲に応じたアートボードごとの入れ物を作る
     * @param {Document} doc - 対象のドキュメント
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @returns {Object} { groups, groupIndexMap, activeArtboardIndex, outsideGroup }（各グループは { name, items }）
     */
    function createArtboardGroups(doc, scopeMode) {
        var artboardGroups = {
            groups: [],
            groupIndexMap: {},
            activeArtboardIndex: doc.artboards.getActiveArtboardIndex(),
            outsideGroup: { name: getLabel('fallbackName.outsideArtboards'), items: [] }
        };

        for (var i = 0; i < doc.artboards.length; i++) {
            if (scopeMode === 'current' && i !== artboardGroups.activeArtboardIndex) continue;
            artboardGroups.groupIndexMap[i] = artboardGroups.groups.length;
            artboardGroups.groups.push({ name: getArtboardExportHeading(doc, i), items: [] });
        }
        return artboardGroups;
    }

    /**
     * 値をアートボードのグループへ振り分ける（範囲外は捨て、ドキュメント全体ならアートボード外へ）
     * @param {Object} artboardGroups - createArtboardGroups() の結果
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @param {number} artboardIndex - 値のあるアートボードの番号（どこにも無ければ -1）
     * @param {Object} exportItem - 振り分ける値（{ text, textFrame }）
     * @returns {void}
     */
    function addToArtboardGroup(artboardGroups, scopeMode, artboardIndex, exportItem) {
        if (artboardGroups.groupIndexMap.hasOwnProperty(artboardIndex)) {
            artboardGroups.groups[artboardGroups.groupIndexMap[artboardIndex]].items.push(exportItem);
        } else if (scopeMode === 'all' && artboardIndex < 0) {
            artboardGroups.outsideGroup.items.push(exportItem);
        }
    }

    /**
     * 書き出すテキスト（テキストフレームとシンボル内テキスト）をアートボードごとにまとめる
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} exportSettings - { includeText, includeFonts, openAfter }
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden, includeSymbols }
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @returns {Object[]} { name, items（{ text, textFrame, fontTriples }） } の配列
     */
    function collectExportGroups(doc, exportSettings, collectOptions, scopeMode) {
        var artboardGroups = createArtboardGroups(doc, scopeMode);
        var textFrames = collectFramesByScope(doc, scopeMode, collectOptions);
        var i;

        for (i = 0; i < textFrames.length; i++) {
            addToArtboardGroup(artboardGroups, scopeMode, getItemArtboardIndex(doc, textFrames[i]),
                { text: textFrames[i].contents, textFrame: textFrames[i], fontTriples: null });
        }
        /* シンボルの展開は重いので、書き出すものがあるときだけ / Expanding symbols is slow, so only when something is exported */
        if ((exportSettings.includeText || exportSettings.includeFonts) && collectOptions.includeSymbols) {
            var symbolEntries = collectSymbolTextsByArtboard(doc, collectOptions, scopeMode);
            for (i = 0; i < symbolEntries.length; i++) {
                addToArtboardGroup(artboardGroups, scopeMode, symbolEntries[i].artboardIndex,
                    { text: symbolEntries[i].text, textFrame: null, fontTriples: symbolEntries[i].fontTriples });
            }
        }

        if (artboardGroups.outsideGroup.items.length > 0) {
            artboardGroups.groups.push(artboardGroups.outsideGroup);
        }
        return artboardGroups.groups;
    }

    /**
     * 書き出すテキストをアートボードごとに組み立てる
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} exportSettings - { includeText, includeFonts, openAfter }
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden, includeSymbols }
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @returns {string} 書き出す内容
     */
    function buildExportText(doc, exportSettings, collectOptions, scopeMode) {
        var exportGroups = collectExportGroups(doc, exportSettings, collectOptions, scopeMode);
        var exportLines = [];

        for (var i = 0; i < exportGroups.length; i++) {
            var groupItems = exportGroups[i].items;
            if (i > 0) exportLines.push('');
            exportLines.push('---' + exportGroups[i].name + '---');

            if (exportSettings.includeText) {
                exportLines.push('[Text]');
                for (var j = 0; j < groupItems.length; j++) {
                    exportLines.push(groupItems[j].text);
                }
                exportLines.push('');
            }

            if (exportSettings.includeFonts) {
                /* テキストフレームは文字から、シンボル内テキストは集めた組から / Frames are read per character, symbols use collected triples */
                var groupFrames = [];
                var tripleLists = [];
                for (var k = 0; k < groupItems.length; k++) {
                    if (groupItems[k].textFrame) groupFrames.push(groupItems[k].textFrame);
                    if (groupItems[k].fontTriples) tripleLists.push(groupItems[k].fontTriples);
                }
                tripleLists.unshift(collectUniqueFontTriples(groupFrames));
                exportLines.push('[Font Names]');
                exportLines = exportLines.concat(buildFontExportLines(mergeFontTriples(tripleLists)));
                exportLines.push('');
            }
        }

        return exportLines.join('\n');
    }

    /**
     * 書き出すファイルのパス（デスクトップの text-<ドキュメント名>-<日時>.txt）を返す
     * @param {Document} doc - 対象のドキュメント
     * @returns {string} ファイルのパス
     */
    function buildExportFilePath(doc) {
        var documentBaseName = doc.name.replace(/\.[^\.]+$/, '');
        return Folder.desktop.fsName + '/text-' + sanitizeFileName(documentBaseName) + '-' + getDateTimeStamp() + '.txt';
    }

    /**
     * UTF-8・LF でテキストファイルを書き出す
     * @param {string} filePath - 書き出すパス
     * @param {string} fileContent - 書き出す内容
     * @returns {File} 書き出したファイル
     */
    function writeTextFile(filePath, fileContent) {
        var exportFile = new File(filePath);
        exportFile.encoding = 'UTF-8';
        exportFile.lineFeed = 'Unix';
        if (!exportFile.open('w')) {
            throw new Error('open failed: ' + filePath);
        }
        exportFile.write(fileContent);
        exportFile.close();
        return exportFile;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ダイアログを作る関数より前）に貼る。
    //    識別子は BUTTON_ROW_* / addButtonRow
    // 2. ダイアログの最後で行を作り、ボタンは btn 接頭辞の変数で左右のグループに足す（キャンセル → OK の順）
    //      var buttonRow = addButtonRow(dialog);
    //      var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel("button.preferences"));
    //      var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    //      var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    //    左右中央に並べるときは addButtonRow(dialog, { centered: true }) にして、buttonRow.rowGroup に直接足す
    // 3. 行の上の余白は BUTTON_ROW_TOP_MARGIN で決める。左右の余白はダイアログの margins に任せる
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ボタン行（再利用パーツ）ここまで / End of the reusable button row
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /**
     * 書き出しオプションのダイアログを表示する
     * @returns {Object|null} OK なら { includeText, includeFonts, openAfter }、キャンセルなら null
     */
    function showExportOptionsDialog() {
        var exportDialog = new Window('dialog', getLabel('dialog.exportOptions'));
        exportDialog.orientation = 'column';
        exportDialog.alignChildren = ['fill', 'top'];

        var exportContentPanel = exportDialog.add('panel', undefined, getLabel('panel.exportContent'));
        exportContentPanel.orientation = 'column';
        exportContentPanel.alignChildren = ['left', 'top'];
        exportContentPanel.margins = EXPORT_PANEL_MARGINS;

        var cbExportText = exportContentPanel.add('checkbox', undefined, getLabel('checkbox.exportIncludeText'));
        cbExportText.value = true;

        var cbExportFonts = exportContentPanel.add('checkbox', undefined, getLabel('checkbox.exportIncludeFonts'));
        cbExportFonts.value = true;

        /* 「書き出し後にファイルを開く」はパネルの外に置く / "Open file after export" sits below the panel */
        var cbOpenAfter = exportDialog.add('checkbox', undefined, getLabel('checkbox.exportOpenAfter'));
        cbOpenAfter.value = true;

        var buttonRow = addButtonRow(exportDialog, { centered: true });
        var btnCancel = buttonRow.rowGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        var btnOK = buttonRow.rowGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });
        exportDialog.defaultElement = btnOK;

        prepareDialogWindow(exportDialog, SCRIPT_NAME + "_exportOptions");
        if (exportDialog.show() !== 1) {
            return null;
        }

        return {
            includeText: cbExportText.value,
            includeFonts: cbExportFonts.value,
            openAfter: cbOpenAfter.value
        };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 右カラムのパネル（縦並び・左揃え）を追加する
     * @param {Group} parentGroup - 追加先
     * @param {string} labelPath - パネル名の LABELS パス
     * @returns {Panel} 追加したパネル
     */
    function addScopePanel(parentGroup, labelPath) {
        var scopePanel = parentGroup.add("panel", undefined, getLabel(labelPath));
        scopePanel.orientation = "column";
        scopePanel.alignChildren = ["left", "top"];
        scopePanel.margins = SCOPE_PANEL_MARGINS;
        return scopePanel;
    }

    /**
     * チェックボックスを追加し、初期値と tooltip を設定する
     * @param {Object} parentContainer - 追加先
     * @param {string} name - LABELS.checkbox のキー（tooltip も同じキーがあれば付ける）
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parentContainer, name, initialValue) {
        var checkbox = parentContainer.add("checkbox", undefined, getLabel("checkbox." + name));
        checkbox.value = initialValue;
        if (LABELS.tooltip.hasOwnProperty(name)) checkbox.helpTip = getLabel("tooltip." + name);
        return checkbox;
    }

    /**
     * 名前一覧用のタブ（見出し＋横並びの切り替え）を追加する
     * @param {TabbedPanel} dialogTabs - 追加先
     * @param {string} tabLabelPath - タブ名の LABELS パス
     * @param {string} headingLabelPath - 見出しの LABELS パス
     * @returns {{tab: Tab, choiceRow: Group}} 追加したタブと、切り替えを並べる行
     */
    function addInfoTab(dialogTabs, tabLabelPath, headingLabelPath) {
        var infoTab = dialogTabs.add("tab", undefined, getLabel(tabLabelPath));
        infoTab.orientation = "column";
        infoTab.alignChildren = ["fill", "top"];
        infoTab.margins = INFO_TAB_MARGINS;
        infoTab.spacing = INFO_TAB_SPACING;

        infoTab.add("statictext", undefined, labelText(headingLabelPath));

        var choiceRow = infoTab.add("group");
        choiceRow.orientation = "row";
        choiceRow.alignChildren = ["left", "center"];
        choiceRow.spacing = CHOICE_ROW_SPACING;
        return { tab: infoTab, choiceRow: choiceRow };
    }

    /**
     * 読み取り専用の複数行テキスト欄を追加する
     * @param {Tab} parentTab - 追加先
     * @returns {EditText} 追加した欄
     */
    function addReadOnlyTextArea(parentTab) {
        var textArea = parentTab.add("edittext", undefined, "", { multiline: true, scrolling: true, readonly: true });
        textArea.alignment = ["fill", "fill"];
        return textArea;
    }

    /**
     * ［テキスト］タブの左カラム（テキスト一覧・編集欄）を組む
     * @param {Group} editColumnsGroup - 追加先
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @returns {void}
     */
    function buildTextColumn(editColumnsGroup, dialogControls) {
        var textColumn = editColumnsGroup.add("group");
        textColumn.orientation = "column";
        textColumn.alignChildren = ["fill", "fill"];

        textColumn.add("statictext", undefined, labelText("fieldLabel.textList"));
        dialogControls.textListBox = textColumn.add("listbox", TEXT_LIST_BOUNDS, []);
        dialogControls.textListBox.helpTip = getLabel("tooltip.textList");

        textColumn.add("statictext", undefined, labelText("fieldLabel.textEdit"));
        dialogControls.textEditBox = textColumn.add("edittext", TEXT_EDIT_BOUNDS, "", { multiline: true, scrolling: true });
        dialogControls.textEditBox.helpTip = getLabel("tooltip.textEdit").replace("{softBreak}", SOFT_BREAK);

        var updateButtonRow = textColumn.add("group");
        updateButtonRow.orientation = "row";
        updateButtonRow.alignment = ["fill", "top"];
        updateButtonRow.alignChildren = ["left", "center"];
        dialogControls.cbSelectedOnly = addCheckbox(updateButtonRow, "selectedOnly", false);

        var spacer = updateButtonRow.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        dialogControls.btnUpdateText = updateButtonRow.add("button", undefined, getLabel("button.updateText"));
        dialogControls.btnUpdateText.helpTip = getLabel("tooltip.updateText");
    }

    /**
     * ［対象範囲］パネル（現在／すべてのアートボード・アートボード外）を組む
     * @param {Group} scopeColumn - 追加先
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @returns {void}
     */
    function addArtboardScopePanel(scopeColumn, dialogControls) {
        var artboardScopePanel = addScopePanel(scopeColumn, "panel.artboardScope");
        dialogControls.rbCurrentArtboard = artboardScopePanel.add("radiobutton", undefined, getLabel("radio.currentArtboard"));
        dialogControls.rbAllArtboards = artboardScopePanel.add("radiobutton", undefined, getLabel("radio.allArtboards"));
        dialogControls.rbCurrentArtboard.value = true;

        dialogControls.cbWholeDocument = addCheckbox(artboardScopePanel, "wholeDocument", false);
        dialogControls.cbWholeDocument.enabled = false;
    }

    /**
     * ［対象テキスト］パネル（//レイヤー・ロック・非表示・シンボル）を組む
     * @param {Group} scopeColumn - 追加先
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @returns {void}
     */
    function addTargetTextPanel(scopeColumn, dialogControls) {
        var targetTextPanel = addScopePanel(scopeColumn, "panel.targetText");
        dialogControls.cbIncludeCommentLayers = addCheckbox(targetTextPanel, "includeCommentLayers", false);
        dialogControls.cbIncludeLocked = addCheckbox(targetTextPanel, "includeLocked", false);
        dialogControls.cbIncludeHidden = addCheckbox(targetTextPanel, "includeHidden", false);
        dialogControls.cbIncludeSymbols = addCheckbox(targetTextPanel, "includeSymbols", DEFAULT_INCLUDE_SYMBOLS);
    }

    /**
     * ［並び順］パネルを組む
     * @param {Group} scopeColumn - 追加先
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @returns {void}
     */
    function addSortPanel(scopeColumn, dialogControls) {
        var sortPanel = addScopePanel(scopeColumn, "panel.sort");
        var sortRadioGroup = sortPanel.add("group");
        sortRadioGroup.orientation = "row";
        sortRadioGroup.alignChildren = ["left", "center"];
        dialogControls.rbSortNone = sortRadioGroup.add("radiobutton", undefined, getLabel("radio.sortNone"));
        dialogControls.rbSortPosition = sortRadioGroup.add("radiobutton", undefined, getLabel("radio.sortPosition"));
        dialogControls.rbSortPosition.helpTip = getLabel("tooltip.sortPosition");
        dialogControls.rbSortAlphabetical = sortRadioGroup.add("radiobutton", undefined, getLabel("radio.sortAlphabetical"));
        dialogControls.rbSortAlphabetical.helpTip = getLabel("tooltip.sortAlphabetical");
        dialogControls.rbSortNone.value = true;
    }

    /**
     * ［オプション］パネル（一括編集・書式を保持）を組む
     * @param {Group} scopeColumn - 追加先
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @returns {void}
     */
    function addEditOptionsPanel(scopeColumn, dialogControls) {
        var editOptionsPanel = addScopePanel(scopeColumn, "panel.editOptions");
        dialogControls.cbMergeDuplicates = addCheckbox(editOptionsPanel, "mergeDuplicates", true);
        dialogControls.cbKeepFormat = addCheckbox(editOptionsPanel, "keepFormat", true);
    }

    /**
     * ［テキスト］タブを組む
     * @param {TabbedPanel} dialogTabs - 追加先
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @returns {Tab} ［テキスト］タブ
     */
    function buildTextEditTab(dialogTabs, dialogControls) {
        var textEditTab = dialogTabs.add("tab", undefined, getLabel("tab.editText"));
        textEditTab.orientation = "row";
        textEditTab.alignChildren = ["fill", "fill"];
        textEditTab.margins = EDIT_TAB_MARGINS;
        textEditTab.spacing = EDIT_TAB_SPACING;

        var editColumnsGroup = textEditTab.add("group");
        editColumnsGroup.orientation = "row";
        editColumnsGroup.alignChildren = ["fill", "fill"];

        buildTextColumn(editColumnsGroup, dialogControls);

        var scopeColumn = editColumnsGroup.add("group");
        scopeColumn.orientation = "column";
        scopeColumn.alignChildren = ["fill", "top"];
        addArtboardScopePanel(scopeColumn, dialogControls);
        addTargetTextPanel(scopeColumn, dialogControls);
        addSortPanel(scopeColumn, dialogControls);
        addEditOptionsPanel(scopeColumn, dialogControls);
        return textEditTab;
    }

    /**
     * ［レイヤー名］［アートボード名］［フォント名］タブを組む
     * @param {TabbedPanel} dialogTabs - 追加先
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @returns {void}
     */
    function buildInfoTabs(dialogTabs, dialogControls) {
        var layerNamesTab = addInfoTab(dialogTabs, "tab.layerNames", "fieldLabel.layerNameList");
        dialogControls.rbLayerScopeTop = layerNamesTab.choiceRow.add("radiobutton", undefined, getLabel("radio.layerScopeTop"));
        dialogControls.rbLayerScopeAll = layerNamesTab.choiceRow.add("radiobutton", undefined, getLabel("radio.layerScopeAll"));
        dialogControls.rbLayerScopeTop.value = true;
        dialogControls.layerNameTextArea = addReadOnlyTextArea(layerNamesTab.tab);

        var artboardNamesTab = addInfoTab(dialogTabs, "tab.artboardNames", "fieldLabel.artboardNameList");
        dialogControls.rbArtboardNameNumbered = artboardNamesTab.choiceRow.add("radiobutton", undefined, getLabel("radio.artboardNameNumbered"));
        dialogControls.rbArtboardNameOnly = artboardNamesTab.choiceRow.add("radiobutton", undefined, getLabel("radio.artboardNameOnly"));
        dialogControls.rbArtboardNameNumbered.value = true;
        dialogControls.artboardNameTextArea = addReadOnlyTextArea(artboardNamesTab.tab);

        var fontNamesTab = addInfoTab(dialogTabs, "tab.fontNames", "fieldLabel.fontList");
        dialogControls.cbFontPS = addCheckbox(fontNamesTab.choiceRow, "fontPS", true);
        dialogControls.cbFontFamily = addCheckbox(fontNamesTab.choiceRow, "fontFamily", true);
        dialogControls.cbFontStyle = addCheckbox(fontNamesTab.choiceRow, "fontStyle", true);

        dialogControls.fontNameListBox = fontNamesTab.tab.add("listbox", undefined, [], {
            numberOfColumns: 3,
            showHeaders: true,
            columnTitles: [getLabel("listColumn.fontPS"), getLabel("listColumn.fontFamily"), getLabel("listColumn.fontStyle")],
            columnWidths: FONT_COLUMN_WIDTHS
        });
        dialogControls.fontNameListBox.alignment = ["fill", "fill"];
        dialogControls.fontNameListBox.helpTip = getLabel("tooltip.fontList");
    }

    /**
     * ボタン行（左に書き出し・コピー、右にキャンセル・OK）を組む
     * @param {Window} textScopeDialog - 追加先のダイアログ
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @returns {void}
     */
    function buildMainButtons(textScopeDialog, dialogControls) {
        var buttonRow = addButtonRow(textScopeDialog);
        dialogControls.btnExportText = buttonRow.leftGroup.add("button", undefined, getLabel("button.exportText"));
        dialogControls.btnExportText.helpTip = getLabel("tooltip.exportText");
        dialogControls.btnCopyText = buttonRow.leftGroup.add("button", undefined, getLabel("button.copyText"));
        dialogControls.btnCopyText.helpTip = getLabel("tooltip.copyText");

        dialogControls.btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        dialogControls.btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        textScopeDialog.defaultElement = dialogControls.btnOK;
    }

    /**
     * メインダイアログを組み立てる（イベントは未接続）
     * @returns {Object} コントロールをまとめたオブジェクト
     */
    function buildDialog() {
        var dialogControls = {};
        var textScopeDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        textScopeDialog.orientation = "column";
        textScopeDialog.alignChildren = ["fill", "top"];
        dialogControls.textScopeDialog = textScopeDialog;

        /* ダイアログ全体をタブで構成 / Use tabs for the whole dialog */
        var dialogTabs = textScopeDialog.add("tabbedpanel");
        dialogTabs.alignChildren = ["fill", "fill"];
        dialogTabs.preferredSize = TABS_SIZE;
        dialogTabs.margins = TABS_MARGINS;

        var textEditTab = buildTextEditTab(dialogTabs, dialogControls);
        buildInfoTabs(dialogTabs, dialogControls);
        dialogTabs.selection = textEditTab;

        buildMainButtons(textScopeDialog, dialogControls);
        return dialogControls;
    }

    // =========================================
    // 一覧の更新 / Updating the lists
    // =========================================

    /**
     * 対象テキストの範囲を返す
     * @param {Object} dialogControls - buildDialog() の結果
     * @returns {string} "all" / "allArtboards" / "current"
     */
    function getCurrentScopeMode(dialogControls) {
        if (dialogControls.cbWholeDocument.value) return 'all';
        if (dialogControls.rbAllArtboards.value) return 'allArtboards';
        return 'current';
    }

    /**
     * 対象テキストの条件をダイアログから読む
     * @param {Object} dialogControls - buildDialog() の結果
     * @returns {Object} { includeComment, includeLocked, includeHidden, includeSymbols }
     */
    function readCollectOptions(dialogControls) {
        return {
            includeComment: dialogControls.cbIncludeCommentLayers.value,
            includeLocked: dialogControls.cbIncludeLocked.value,
            includeHidden: dialogControls.cbIncludeHidden.value,
            includeSymbols: dialogControls.cbIncludeSymbols.value
        };
    }

    /**
     * シンボル内のテキストを集め直す（［シンボル内のテキスト］がオフなら空）
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function refreshSymbolEntries(editSession) {
        var dialogControls = editSession.dialogControls;
        var collectOptions = readCollectOptions(dialogControls);
        editSession.symbolEntries = collectOptions.includeSymbols
            ? collectSymbolTextEntries(editSession.doc, collectOptions, getCurrentScopeMode(dialogControls))
            : [];
    }

    /**
     * 一覧に並べるシンボル内テキストを返す（ABC順のときは内容の順に並べ替える）
     * @param {Object} editSession - 編集の状態
     * @returns {Object[]} readSymbolTextEntries() の要素の配列
     */
    function getListedSymbolEntries(editSession) {
        var symbolEntries = editSession.symbolEntries.slice(0);
        if (editSession.dialogControls.rbSortAlphabetical.value) {
            symbolEntries.sort(function (firstEntry, secondEntry) {
                var firstText = firstEntry.contents.toLowerCase();
                var secondText = secondEntry.contents.toLowerCase();
                if (firstText < secondText) return -1;
                if (firstText > secondText) return 1;
                return 0;
            });
        }
        return symbolEntries;
    }

    /**
     * テキスト一覧を並べ直す（テキストフレームは集め直し、シンボル内テキストは集めた結果を末尾に並べる）
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function refreshTextList(editSession) {
        var dialogControls = editSession.dialogControls;
        var textFrames = collectFramesByScope(editSession.doc, getCurrentScopeMode(dialogControls), readCollectOptions(dialogControls));
        if (dialogControls.rbSortPosition.value) {
            sortByPosition(textFrames);
        } else if (dialogControls.rbSortAlphabetical.value) {
            sortByContent(textFrames);
        }

        editSession.duplicateMap = [];
        if (dialogControls.cbMergeDuplicates.value) {
            var duplicateGroups = groupDuplicateFrames(textFrames);
            editSession.duplicateMap = duplicateGroups.duplicateMap;
            textFrames = duplicateGroups.uniqueFrames;
        }
        editSession.textFrameList = textFrames;

        /* テキストフレームの行は番号、シンボル内テキストの行は ♣ で始める / Frame rows are numbered, symbol rows start with ♣ */
        var listRows = [];
        var rowLabels = [];
        var i;
        for (i = 0; i < textFrames.length; i++) {
            listRows.push({ kind: 'frame', index: i, contents: textFrames[i].contents });
            rowLabels.push((i + 1) + getLabel("format.itemPrefix") + buildListLabel(textFrames[i].contents));
        }
        var symbolEntries = getListedSymbolEntries(editSession);
        for (i = 0; i < symbolEntries.length; i++) {
            listRows.push({ kind: 'symbol', entry: symbolEntries[i], contents: symbolEntries[i].contents });
            rowLabels.push(getLabel("format.symbolRowMark") + buildListLabel(symbolEntries[i].contents)
                + getLabel("format.symbolRowSuffix").split('{symbolName}').join(symbolEntries[i].symbolName));
        }
        editSession.listRows = listRows;

        var textListBox = dialogControls.textListBox;
        textListBox.removeAll();
        for (i = 0; i < rowLabels.length; i++) {
            textListBox.add("item", rowLabels[i]);
        }
        if (listRows.length > 0) {
            textListBox.selection = 0;
            setEditTarget(editSession, listRows[0]);
        } else {
            setEditTarget(editSession, null);
        }
    }

    /**
     * ［レイヤー名］タブを更新する
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function refreshLayerNameList(editSession) {
        var dialogControls = editSession.dialogControls;
        var layerNames = collectDocumentLayerNames(editSession.doc, dialogControls.rbLayerScopeAll.value);
        layerNames.sort();
        dialogControls.layerNameTextArea.text = layerNames.join("\n");
    }

    /**
     * ［アートボード名］タブを更新する
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function refreshArtboardNameList(editSession) {
        var doc = editSession.doc;
        var dialogControls = editSession.dialogControls;
        var artboardNames = [];
        for (var i = 0; i < doc.artboards.length; i++) {
            artboardNames.push(dialogControls.rbArtboardNameOnly.value ? getArtboardName(doc, i) : getNumberedArtboardName(doc, i));
        }
        artboardNames.sort();
        dialogControls.artboardNameTextArea.text = artboardNames.join("\n");
    }

    /**
     * ［フォント名］タブを更新する（フォントの集計は初回と、シンボルを書き換えたあとだけ行う）
     * ドキュメント全体のテキストと、配置されているシンボル内のテキストを対象にする
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function refreshFontNameList(editSession) {
        var dialogControls = editSession.dialogControls;
        if (!editSession.fontRows) {
            var doc = editSession.doc;
            editSession.allSymbolEntries = collectSymbolTextEntries(doc, {
                includeComment: true,
                includeLocked: true,
                includeHidden: true
            }, 'all');
            var tripleLists = [collectUniqueFontTriples(collectAllTextFrames(doc))];
            for (var j = 0; j < editSession.allSymbolEntries.length; j++) {
                tripleLists.push(editSession.allSymbolEntries[j].fontTriples);
            }
            editSession.fontRows = mergeFontTriples(tripleLists);
            editSession.fontRows.sort(function (firstRow, secondRow) {
                var firstKey = getFontTripleKey(firstRow);
                var secondKey = getFontTripleKey(secondRow);
                if (firstKey < secondKey) return -1;
                if (firstKey > secondKey) return 1;
                return 0;
            });
        }

        var fontRows = editSession.fontRows;
        var fontNameListBox = dialogControls.fontNameListBox;
        fontNameListBox.removeAll();
        for (var i = 0; i < fontRows.length; i++) {
            var fontListRow = fontNameListBox.add('item', dialogControls.cbFontPS.value ? fontRows[i].psName : '');
            fontListRow.subItems[0].text = dialogControls.cbFontFamily.value ? fontRows[i].familyName : '';
            fontListRow.subItems[1].text = dialogControls.cbFontStyle.value ? fontRows[i].styleName : '';
            fontListRow.psNameKey = fontRows[i].psName;
        }
    }

    // =========================================
    // 編集 / Editing
    // =========================================

    /**
     * 選択していたテキストフレームだけを残す
     * @param {TextFrame[]} textFrames - 編集対象のテキストフレーム
     * @param {TextFrame[]} selectedFrames - 選択していたテキストフレーム
     * @returns {TextFrame[]} 選択していたものだけの配列
     */
    function filterSelectedFrames(textFrames, selectedFrames) {
        var filteredFrames = [];
        for (var i = 0; i < textFrames.length; i++) {
            if (indexOfItem(selectedFrames, textFrames[i]) >= 0) filteredFrames.push(textFrames[i]);
        }
        return filteredFrames;
    }

    /**
     * シンボルのインスタンスを選択していたかを返す（書き換えでシンボルが差し替わるので、その都度インスタンスから引く）
     * @param {Symbol} symbol - 調べるシンボル
     * @param {SymbolItem[]} selectedSymbolItems - 選択していたシンボルインスタンス
     * @returns {boolean} 選択していたインスタンスがあれば true
     */
    function isSymbolSelected(symbol, selectedSymbolItems) {
        for (var i = 0; i < selectedSymbolItems.length; i++) {
            try {
                if (selectedSymbolItems[i].symbol === symbol) return true;
            } catch (removedError) { }
        }
        return false;
    }

    /**
     * 編集の対象を切り替え、編集欄にその内容を出す
     * @param {Object} editSession - 編集の状態
     * @param {Object|null} listRow - editSession.listRows の要素（{ kind: "frame" | "symbol", index | entry, contents }）または null
     * @returns {void}
     */
    function setEditTarget(editSession, listRow) {
        editSession.editTarget = listRow;
        editSession.dialogControls.textEditBox.text = listRow ? toEditText(listRow.contents) : "";
    }

    /**
     * 選択中の行のテキストに編集を反映する（まとめているときは同じ内容の全フレームへ。
     * ［選択しているテキストのみ］のときは選択していたものだけ。変えていなければ何もしない）
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function applyCurrentEdit(editSession) {
        var editTarget = editSession.editTarget;
        if (!editTarget) return;
        var dialogControls = editSession.dialogControls;
        var newText = toFrameText(dialogControls.textEditBox.text);
        if (newText === editTarget.contents) return;
        var keepFormat = dialogControls.cbKeepFormat.value;
        var selectedOnly = dialogControls.cbSelectedOnly.value;

        if (editTarget.kind === 'symbol') {
            if (selectedOnly && !isSymbolSelected(editTarget.entry.symbol, editSession.selectedItems.symbolItems)) return;
            if (!replaceSymbolText(editSession.doc, editTarget.entry, newText, keepFormat)) {
                alert(getLabel('alert.symbolUpdateFailed'));
            }
            return;
        }

        var targetFrames = editSession.duplicateMap[editTarget.index] || [editSession.textFrameList[editTarget.index]];
        if (selectedOnly) targetFrames = filterSelectedFrames(targetFrames, editSession.selectedItems.frames);
        for (var i = 0; i < targetFrames.length; i++) {
            replaceTextContents(targetFrames[i], newText, keepFormat);
        }
    }

    /**
     * 編集をドキュメントに反映し、一覧を並べ直して同じ位置の行を選び直す
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function applyEditAndRefresh(editSession) {
        var textListBox = editSession.dialogControls.textListBox;
        var selectedIndex = textListBox.selection ? textListBox.selection.index : 0;
        var isSymbolEdit = editSession.editTarget && editSession.editTarget.kind === 'symbol';

        applyCurrentEdit(editSession);
        /* シンボルは定義ごと差し替わるので集め直す / The symbol was swapped, so collect symbols again */
        if (isSymbolEdit) {
            refreshSymbolEntries(editSession);
            editSession.fontRows = null;
            refreshFontNameList(editSession);
        }
        refreshTextList(editSession);
        app.redraw();

        var rowCount = editSession.listRows.length;
        if (rowCount > 0) {
            var rowIndex = Math.min(selectedIndex, rowCount - 1);
            textListBox.selection = rowIndex;
            setEditTarget(editSession, editSession.listRows[rowIndex]);
        }
    }

    /**
     * 書き出しオプションを尋ね、デスクトップにテキストを書き出す
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function exportTextToDesktop(editSession) {
        var doc = editSession.doc;
        var dialogControls = editSession.dialogControls;
        try {
            var exportSettings = showExportOptionsDialog();
            if (!exportSettings) return;
            var exportContent = buildExportText(doc, exportSettings, readCollectOptions(dialogControls), getCurrentScopeMode(dialogControls));
            var exportFile = writeTextFile(buildExportFilePath(doc), exportContent);
            if (exportSettings.openAfter) exportFile.execute();
        } catch (e) {
            alert(getLabel('alert.exportFailed') + '\n' + e);
        }
    }

    /**
     * テキスト一覧に並んでいるテキストを、1行に1つずつクリップボードへコピーする
     * 編集欄に一時的に入れて選択し、メニューの［コピー］を実行する（編集中の内容はあとで戻す）
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function copyListToClipboard(editSession) {
        var doc = editSession.doc;
        var textEditBox = editSession.dialogControls.textEditBox;
        var copyLines = [];
        for (var i = 0; i < editSession.listRows.length; i++) {
            /* 段落の区切りと強制改行はどちらも改行にする / Paragraph and forced line breaks both become newlines */
            copyLines.push(editSession.listRows[i].contents.replace(/[\r\u0003]/g, '\n'));
        }
        var copyText = copyLines.join('\n');

        var savedEditText = textEditBox.text;
        /* オブジェクトがコピーされないよう選択を外しておく / Deselect so objects are not copied instead */
        var savedSelection = toItemArray(doc.selection);
        try {
            doc.selection = null;
            textEditBox.text = copyText;
            textEditBox.active = true;
            textEditBox.textselection = copyText;
            app.executeMenuCommand('copy');
        } catch (e) {
            alert(e);
        } finally {
            textEditBox.text = savedEditText;
            for (var j = 0; j < savedSelection.length; j++) {
                /* 消えたアイテムやロックされたアイテムは選び直せない / Removed or locked items cannot be reselected */
                try {
                    savedSelection[j].selected = true;
                } catch (reselectError) { }
            }
        }
    }

    // =========================================
    // イベントハンドラ / Event handlers
    // =========================================

    /**
     * 一覧・編集欄・編集オプションのイベントを接続する
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function bindEditEvents(editSession) {
        var dialogControls = editSession.dialogControls;

        /* 一覧の選択で編集欄を更新 / Update the edit box when the list selection changes */
        dialogControls.textListBox.onChange = function () {
            if (dialogControls.textListBox.selection === null) return;
            setEditTarget(editSession, editSession.listRows[dialogControls.textListBox.selection.index]);
        };

        /* Shift+Enter でソフト改行文字を挿入 / Insert the soft-break placeholder with Shift+Enter */
        dialogControls.textEditBox.addEventListener('keydown', function (keyEvent) {
            if (ScriptUI.environment.keyboardState.shiftKey && keyEvent.keyName === 'Enter') {
                this.textselection = SOFT_BREAK;
                keyEvent.preventDefault();
            }
        });
        dialogControls.cbMergeDuplicates.onClick = function () { refreshTextList(editSession); };
    }

    /**
     * 対象範囲・対象テキスト・並び順のイベントを接続する
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function bindScopeEvents(editSession) {
        var dialogControls = editSession.dialogControls;

        /**
         * シンボル内のテキストも集め直してテキスト一覧を並べ直す
         * @returns {void}
         */
        function refreshAllLists() {
            refreshSymbolEntries(editSession);
            refreshTextList(editSession);
        }

        /**
         * テキスト一覧だけを並べ直す
         * @returns {void}
         */
        function refreshTextListOnly() {
            refreshTextList(editSession);
        }

        dialogControls.rbCurrentArtboard.onClick = function () {
            dialogControls.cbWholeDocument.enabled = false;
            dialogControls.cbWholeDocument.value = false;
            refreshAllLists();
        };
        dialogControls.rbAllArtboards.onClick = function () {
            dialogControls.cbWholeDocument.enabled = true;
            refreshAllLists();
        };
        dialogControls.cbWholeDocument.onClick = refreshAllLists;
        dialogControls.cbIncludeCommentLayers.onClick = refreshAllLists;
        dialogControls.cbIncludeLocked.onClick = refreshAllLists;
        dialogControls.cbIncludeHidden.onClick = refreshAllLists;
        dialogControls.cbIncludeSymbols.onClick = refreshAllLists;
        dialogControls.rbSortNone.onClick = refreshTextListOnly;
        dialogControls.rbSortPosition.onClick = refreshTextListOnly;
        dialogControls.rbSortAlphabetical.onClick = refreshTextListOnly;
    }

    /**
     * 名前一覧タブのイベントを接続する
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function bindInfoTabEvents(editSession) {
        var dialogControls = editSession.dialogControls;

        /**
         * ［フォント名］タブを更新する
         * @returns {void}
         */
        function refreshFontNames() {
            refreshFontNameList(editSession);
        }

        dialogControls.rbLayerScopeAll.onClick = function () { refreshLayerNameList(editSession); };
        dialogControls.rbLayerScopeTop.onClick = function () { refreshLayerNameList(editSession); };
        dialogControls.rbArtboardNameNumbered.onClick = function () { refreshArtboardNameList(editSession); };
        dialogControls.rbArtboardNameOnly.onClick = function () { refreshArtboardNameList(editSession); };
        dialogControls.cbFontPS.onClick = refreshFontNames;
        dialogControls.cbFontFamily.onClick = refreshFontNames;
        dialogControls.cbFontStyle.onClick = refreshFontNames;

        dialogControls.fontNameListBox.onChange = function () {
            if (dialogControls.fontNameListBox.selection !== null) {
                selectItemsByFont(editSession.doc, dialogControls.fontNameListBox.selection.psNameKey, editSession.allSymbolEntries);
            }
        };
    }

    /**
     * ボタンとダイアログを閉じるときのイベントを接続する
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function bindButtonEvents(editSession) {
        var dialogControls = editSession.dialogControls;

        dialogControls.btnExportText.onClick = function () { exportTextToDesktop(editSession); };
        dialogControls.btnCopyText.onClick = function () { copyListToClipboard(editSession); };
        dialogControls.btnUpdateText.onClick = function () { applyEditAndRefresh(editSession); };

        dialogControls.btnCancel.onClick = function () {
            dialogControls.textScopeDialog.close();
        };

        /* OKボタンで現在の編集を反映して閉じる / Apply the current edit and close when OK is pressed */
        dialogControls.btnOK.onClick = function () {
            applyCurrentEdit(editSession);
            dialogControls.textScopeDialog.close();
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを開き、テキストの一覧と編集を扱う
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        /* 編集の状態 / Edit session state */
        var editSession = {
            doc: app.activeDocument,
            dialogControls: null,
            selectedItems: null,      /* 起動時に選択していたテキストとシンボルインスタンス / Text and symbol instances selected at launch */
            textFrameList: [],        /* 一覧に並ぶテキストフレーム / Text frames listed */
            duplicateMap: [],         /* 行ごとの同じ内容の全フレーム / All frames with the same contents per row */
            symbolEntries: [],        /* シンボル内のテキスト / Text in symbols */
            listRows: [],             /* テキスト一覧の行ごとの編集対象 / Edit target per list row */
            fontRows: null,           /* フォント名タブの行（初回に集計）/ Font tab rows (collected once) */
            allSymbolEntries: [],     /* フォント名タブ用の、ドキュメント全体のシンボル内テキスト / Symbol text across the document for the font tab */
            editTarget: null          /* 編集の対象（listRows の要素）/ Current edit target (an element of listRows) */
        };
        editSession.selectedItems = collectSelectedTextItems(editSession.doc);
        editSession.dialogControls = buildDialog();
        /* 選択が無ければ選べない。あれば最初からオン / Unavailable without a selection; on from the start when there is one */
        var hasSelectedText = editSession.selectedItems.frames.length > 0 || editSession.selectedItems.symbolItems.length > 0;
        editSession.dialogControls.cbSelectedOnly.value = hasSelectedText;
        editSession.dialogControls.cbSelectedOnly.enabled = hasSelectedText;
        bindEditEvents(editSession);
        bindScopeEvents(editSession);
        bindInfoTabEvents(editSession);
        bindButtonEvents(editSession);

        refreshSymbolEntries(editSession);
        refreshTextList(editSession);
        refreshLayerNameList(editSession);
        refreshArtboardNameList(editSession);
        refreshFontNameList(editSession);

        prepareDialogWindow(editSession.dialogControls.textScopeDialog, SCRIPT_NAME);
        editSession.dialogControls.textScopeDialog.show();
    }

    main();

})();
