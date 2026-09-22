#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストを、ダイアログで指定したルールに従って整形します。
行頭・行末スペースの削除、ナンバリングの削除と振り直し、改行の変換、空行の整理などに対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextNormalize.md

### Overview

Tidies the selected text according to the rules you set in the dialog.
It can trim leading and trailing spaces, remove and renumber list numbering, convert returns, and collapse blank lines.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextNormalize.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextNormalize";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextNormalize.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextNormalize.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_WIDTH = 500;                    /* ダイアログとタブの幅 / width of the dialog and the tabs */
    var PANEL_MARGINS = [15, 20, 15, 10];      /* タブ・パネルの余白 [左, 上, 右, 下] / tab and panel margins */
    var SMALL_BUTTON_HEIGHT = 22;              /* 少し小さめのボタンの高さ / height of the compact buttons */
    var CASE_BUTTON_WIDTH = 220;               /* ケース変換ボタンの幅 / width of the case buttons */
    var CASE_EXAMPLE_WIDTH = 240;              /* 変換例の表示幅 / width of the case examples */
    var CASE_EXAMPLE_MAX_LENGTH = 40;          /* 変換例の最大文字数（長いとUIが伸びる） / max length of a case example */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI の表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "テキストのクリーンアップと変換", en: "Clean Up and Convert Text" }
        },
        tab: {
            remove: { ja: "削除", en: "Remove" },
            numbering: { ja: "ナンバリング", en: "Numbering" },
            lines: { ja: "行", en: "Lines" },
            letterCase: { ja: "アルファベット", en: "Letter Case" },
            other: { ja: "その他", en: "Other" }
        },
        panel: {
            numberStyle: { ja: "形式", en: "Format" },
            sort: { ja: "ソート", en: "Sort" }
        },
        checkbox: {
            leadingSpace: { ja: "行頭のスペース", en: "Leading spaces" },
            trailingSpace: { ja: "行末のスペース", en: "Trailing spaces" },
            multipleSpace: { ja: "連続するスペース", en: "Repeated spaces" },
            tabToSpace: { ja: "タブ→半角スペースに", en: "Tabs to spaces" },
            zenkakuToHankaku: { ja: "全角英数字を半角に", en: "Full-width alphanumerics to half-width" },
            hyphenToUnderscore: { ja: "ハイフンをアンダースコアに", en: "Hyphens to underscores" },
            underscoreToHyphen: { ja: "アンダースコアをハイフンに", en: "Underscores to hyphens" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        radio: {
            styleNumDot: { ja: "1. いちご", en: "1. Apple" },
            styleAlphaDot: { ja: "A. いちご", en: "A. Apple" },
            styleCircled: { ja: "\u2460 いちご", en: "\u2460 Apple" },
            styleDot: { ja: "・いちご", en: "\u2022 Apple" },
            styleHyphen: { ja: "- いちご", en: "- Apple" },
            sortAsc: { ja: "ソート", en: "Sort ascending" },
            reverse: { ja: "行を逆順に", en: "Reverse line order" },
            uniqueAdjacent: { ja: "隣接する重複行を削除", en: "Remove adjacent duplicates" }
        },
        button: {
            removeMarker: { ja: "行頭マーカー削除", en: "Remove line markers" },
            renumber: { ja: "ナンバリングの振り直し", en: "Renumber" },
            reset: { ja: "リセット", en: "Reset" },
            forcedToPara: { ja: "強制改行を改行に", en: "Line breaks to paragraph breaks" },
            paraToForced: { ja: "改行を強制改行に", en: "Paragraph breaks to line breaks" },
            compressBlank: { ja: "空行の整理（連続改行の圧縮）", en: "Collapse blank lines" },
            caseUpper: { ja: "すべて大文字に", en: "UPPERCASE" },
            caseLower: { ja: "すべて小文字に", en: "lowercase" },
            caseWord: { ja: "単語の先頭のみ大文字", en: "Capitalize Each Word" },
            caseSentence: { ja: "文頭のみ大文字", en: "Sentence case" },
            caseTitle: { ja: "英語タイトル形式", en: "Title Case" },
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            leadingSpace: { ja: "各行の先頭にある半角・全角スペースを削除します。", en: "Removes spaces at the start of each line." },
            trailingSpace: { ja: "各行の末尾にある半角・全角スペースを削除します。", en: "Removes spaces at the end of each line." },
            multipleSpace: { ja: "2つ以上続くスペースを1つにまとめます。", en: "Collapses runs of spaces into a single space." },
            removeMarker: { ja: "「1.」「・」などの行頭マーカーを削除します。", en: "Removes line markers such as \u00221.\u0022 or \u0022\u2022\u0022." },
            renumber: {
                ja: "行頭マーカーを、右で選んだ形式の連番に付け替えます。",
                en: "Replaces the line markers with a sequence in the format chosen on the right."
            },
            resetNumbering: {
                ja: "ナンバリングの設定を解除し、テキストを開いた時点に戻します。",
                en: "Clears the numbering settings and restores the text as it was when the dialog opened."
            },
            numberStyle: { ja: "振り直したときの連番の形式です。", en: "Format used when renumbering." },
            forcedToPara: { ja: "Shift+Returnの改行を、段落の改行に変えます。", en: "Converts Shift+Return line breaks into paragraph breaks." },
            paraToForced: { ja: "段落の改行を、Shift+Returnの改行に変えます。", en: "Converts paragraph breaks into Shift+Return line breaks." },
            compressBlank: { ja: "2行以上続く空行を1行にまとめます。", en: "Collapses runs of blank lines into one." },
            sortAsc: { ja: "行を昇順に並べ替えます。", en: "Sorts the lines in ascending order." },
            reverse: { ja: "行の並びを逆さにします。", en: "Reverses the order of the lines." },
            uniqueAdjacent: {
                ja: "同じ内容が続く行を1行にまとめます。離れた重複は残ります。",
                en: "Merges consecutive identical lines. Duplicates further apart are kept."
            },
            caseButton: { ja: "右に変換後の例を表示します。", en: "The result is previewed on the right." },
            resetCase: {
                ja: "すべての設定を解除し、テキストを開いた時点に戻します。",
                en: "Clears every setting and restores the text as it was when the dialog opened."
            },
            tabToSpace: { ja: "タブ文字を半角スペース1つに置き換えます。", en: "Replaces each tab with a single space." },
            zenkakuToHankaku: { ja: "全角の英数字を半角に変えます。", en: "Converts full-width alphanumerics to half-width." },
            hyphenToUnderscore: { ja: "ハイフンをアンダースコアに置き換えます。", en: "Replaces hyphens with underscores." },
            underscoreToHyphen: { ja: "アンダースコアをハイフンに置き換えます。", en: "Replaces underscores with hyphens." },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
                en: "Shows the result on the canvas. Cancel restores the original state."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noOption: { ja: "実行する処理を1つ以上チェックしてください。", en: "Select at least one operation to run." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    // =========================================
    // 文字列の変換 / Text conversions
    // =========================================

    /* 半角/全角数字（ナンバリング用） / Half- and full-width digits for numbering */
    var DIGITS_FOR_RENUMBER = "0-9\uFF10-\uFF19";

    /**
     * 連番を A, B, … Z, AA, AB … の英字にする
     * @param {number} sequenceNumber - 1 から始まる番号
     * @returns {string} 英字の番号
     */
    function toAlphabetLabel(sequenceNumber) {
        var remaining = sequenceNumber;
        var letters = "";
        while (remaining > 0) {
            remaining--; /* 1-based */
            letters = String.fromCharCode(65 + (remaining % 26)) + letters;
            remaining = Math.floor(remaining / 26);
        }
        return letters;
    }

    /**
     * 振り直し後の行頭マーカーを作る
     * @param {number} sequenceNumber - 1 から始まる番号
     * @param {string} renumberStyle - "num"=1. / "alpha"=A. / "circled"=① / "dot"=・ / "hyphen"=-
     * @returns {string} 行頭に付ける文字列
     */
    function formatNumberPrefix(sequenceNumber, renumberStyle) {
        if (renumberStyle === "alpha") return toAlphabetLabel(sequenceNumber) + ". ";
        if (renumberStyle === "circled") {
            /* ①(1)〜⑳(20) まで対応。超えたら (n) / Circled up to 20, then (n) */
            if (sequenceNumber >= 1 && sequenceNumber <= 20) {
                return String.fromCharCode(0x2460 + (sequenceNumber - 1)) + " ";
            }
            return "(" + String(sequenceNumber) + ") ";
        }
        if (renumberStyle === "dot") return "・";
        if (renumberStyle === "hyphen") return "- ";
        /* 既定: 1. / default: 1. */
        return String(sequenceNumber) + ". ";
    }

    /**
     * 行頭の番号・記号を外し、指定の形式で連番を振り直す（空の行には振らない）
     * @param {string} sourceText - 対象の文字列
     * @param {string} renumberStyle - 連番の形式（formatNumberPrefix() を参照）
     * @returns {string} 振り直した文字列
     */
    function renumberLineHeads(sourceText, renumberStyle) {
        /* 改行区切り（CR/ETX/LF）を保持して処理する / Keep the separators (CR/ETX/LF) */
        var parts = sourceText.split(/(\r|\u0003|\n)/);
        var sequenceNumber = 1;

        for (var i = 0; i < parts.length; i += 2) {
            var line = parts[i];
            if (line == null) continue;

            /* 行頭インデント（半角/全角スペース） / Leading indent (half- or full-width spaces) */
            var indentMatch = line.match(/^[ \u3000]*/);
            var indent = indentMatch ? indentMatch[0] : "";
            var lineBody = line.substring(indent.length);

            /* 既存の先頭マーカーを除去（番号/ドット/中黒/ハイフン等） / Strip an existing marker */
            lineBody = lineBody.replace(new RegExp(
                "^(?:" +
                "[" + DIGITS_FOR_RENUMBER + "]+(?:[\\.\\uFF0E])?" +
                "|[A-Za-z](?:[\\.\\uFF0E])" +
                "|[\\u2460-\\u2473]" +
                "|・" +
                "|-" +
                ")[ \\u3000]*"), "");

            /* 残りが空（空白のみ含む）なら、番号は振らずにそのまま / Leave blank lines unnumbered */
            if (/^[ \u3000\t]*$/.test(lineBody)) {
                parts[i] = indent + lineBody;
                continue;
            }

            /* 先頭の余計な空白は落として、出力形式を付ける / Drop leading blanks, then prefix */
            lineBody = lineBody.replace(/^[ \u3000\t]+/, "");
            parts[i] = indent + formatNumberPrefix(sequenceNumber, renumberStyle) + lineBody;
            sequenceNumber++;
        }

        return parts.join("");
    }

    /**
     * 全角英数字を半角にする（全角 0-9 / A-Z / a-z）
     * @param {string} sourceText - 対象の文字列
     * @returns {string} 変換した文字列
     */
    function zenkakuAlnumToHankaku(sourceText) {
        var converted = "";
        for (var i = 0; i < sourceText.length; i++) {
            var charCode = sourceText.charCodeAt(i);
            var isZenkakuDigit = (charCode >= 0xFF10 && charCode <= 0xFF19);
            var isZenkakuUpper = (charCode >= 0xFF21 && charCode <= 0xFF3A);
            var isZenkakuLower = (charCode >= 0xFF41 && charCode <= 0xFF5A);
            if (isZenkakuDigit || isZenkakuUpper || isZenkakuLower) {
                converted += String.fromCharCode(charCode - 0xFEE0);
            } else {
                converted += sourceText.charAt(i);
            }
        }
        return converted;
    }

    /**
     * 各単語の先頭の小文字を大文字にする
     * @param {string} sourceText - 対象の文字列
     * @returns {string} 変換した文字列
     */
    function toWordCap(sourceText) {
        return sourceText.replace(/\b([a-z])/g, function (_, firstChar) { return firstChar.toUpperCase(); });
    }

    /**
     * 文頭だけを大文字にする。NASA のような全大文字語と iPhone のような途中に大文字を含む語は保つ
     * @param {string} sourceText - 対象の文字列
     * @returns {string} 変換した文字列
     */
    function toSentenceCasePreserveAcronyms(sourceText) {
        var converted = sourceText;
        /* 1) 保護する語を退避してから全文を小文字化 / Stash protected words, then lowercase everything */
        var placeholders = [];
        function stashWord(word) {
            /* toLowerCase() の影響を受けないよう、英字を含まない Private Use Area のマーカーにする
               Use Private Use Area markers without letters so toLowerCase() leaves them alone */
            var marker = "\uE000" + placeholders.length + "\uE001";
            placeholders.push(word);
            return marker;
        }

        /* 全大文字語（2文字以上） / All-caps words (2+ letters) */
        converted = converted.replace(/\b[A-Z]{2,}\b/g, stashWord);
        /* mixedCase / camelCase（例: iPhone, eBay, PowerPoint, macOS, ChatGPT） */
        converted = converted.replace(/\b[A-Za-z]*[a-z][A-Za-z]*[A-Z][A-Za-z]*\b/g, stashWord);

        converted = converted.toLowerCase();

        /* 2) 文頭および .!? の後の英字を大文字に / Capitalize after sentence starts */
        converted = converted.replace(/(^|[\.\!\?]\s+|[\r\n\u0003]+)([a-z])/g,
            function (_, prefix, firstChar) { return prefix + firstChar.toUpperCase(); });

        /* 3) 退避した語を復元（同じマーカーが何度出ても戻す） / Restore every stashed word */
        for (var i = 0; i < placeholders.length; i++) {
            var marker = "\uE000" + i + "\uE001";
            while (converted.indexOf(marker) !== -1) {
                converted = converted.replace(marker, placeholders[i]);
            }
        }

        return converted;
    }

    /**
     * 英語のタイトル形式にする（John Gruber / John Resig の Title Caps を元にしたロジック。英字にのみ作用）
     * @param {string} sourceText - 対象の文字列
     * @returns {string} 変換した文字列
     */
    function toTitleCase(sourceText) {
        /* 小文字にする語（冠詞・接続詞・前置詞など） / Words kept lowercase (articles, conjunctions, prepositions) */
        var smallWords = "(a|abaft|aboard|about|above|absent|across|afore|after|against|along|alonside|amid|amidst|among|amongst|an|and|apopos|around|as|aside|astride|at|athwart|atop|barring|before|behind|below|beneath|beside|besides|between|betwixt|beyond|but|by|circa|concerning|despite|down|during|except|excluding|failing|following|for|from|given|in|including|inside|into|lest|like|mid|midst|minus|modula|near|next|nor|notwithstanding|of|off|on|onto|oppostie|or|out|outside|over|pace|per|plus|pro|qua|regarding|round|sans|save|than|that|the|through|throughout|till|times|to|toward|towards|under|underneath|unlike|until|unto|up|upon|versus|via|vice|with|within|without|worth|v[.]?|via|vs[.]?)";
        var punctuation = "([!\"#$%&'()*+,./:;<=>?@[\\\\\\]^_`{|}~-]*)";

        function toLower(word) { return word.toLowerCase(); }
        function capitalizeFirst(word) { return word.substr(0, 1).toUpperCase() + word.substr(1); }

        var parts = [];
        var separatorPattern = /[:.;?!] |(?: |^)[\"\u00D2]/g;
        var index = 0;

        while (true) {
            var separatorMatch = separatorPattern.exec(sourceText);

            parts.push(
                sourceText.substring(index, separatorMatch ? separatorMatch.index : sourceText.length)
                    .replace(/\b([A-Za-z][a-z.'\u00D5]*)\b/g, function (all) {
                        return /[A-Za-z]\.[A-Za-z]/.test(all) ? all : capitalizeFirst(all);
                    })
                    .replace(RegExp("\\b" + smallWords + "\\b", "ig"), toLower)
                    .replace(RegExp("^" + punctuation + smallWords + "\\b", "ig"), function (all, leadingPunctuation, word) {
                        return leadingPunctuation + capitalizeFirst(word);
                    })
                    .replace(RegExp("\\b" + smallWords + punctuation + "$", "ig"), capitalizeFirst)
            );

            index = separatorPattern.lastIndex;

            if (separatorMatch) parts.push(separatorMatch[0]);
            else break;
        }

        return parts.join("")
            .replace(/ V(s?)\. /ig, " v$1. ")
            .replace(/(['\u00D5])S\b/ig, "$1s")
            .replace(/\b(AT&T|Q&A)\b/ig, function (all) { return all.toUpperCase(); });
    }

    /**
     * 英字のケースを変換する
     * @param {string} sourceText - 対象の文字列
     * @param {string|null} caseMode - "upper" / "lower" / "word" / "sentence" / "title"（null なら変換しない）
     * @returns {string} 変換した文字列
     */
    function convertCase(sourceText, caseMode) {
        if (caseMode === "upper") return sourceText.toUpperCase();
        if (caseMode === "lower") return sourceText.toLowerCase();
        if (caseMode === "word") return toWordCap(sourceText);
        if (caseMode === "sentence") return toSentenceCasePreserveAcronyms(sourceText);
        if (caseMode === "title") return toTitleCase(sourceText);
        return sourceText;
    }

    /**
     * 段落（\r）で区切った行を並べ替える
     * @param {string} sourceText - 対象の文字列
     * @param {string} lineOrder - "sortAsc"（昇順）/ "reverse"（逆順）/ "uniqueAdjacent"（隣接する重複を削除）
     * @returns {string} 並べ替えた文字列
     */
    function reorderLines(sourceText, lineOrder) {
        var lines = sourceText.split("\r");
        if (lineOrder === "sortAsc") {
            lines.sort();
        } else if (lineOrder === "reverse") {
            lines.reverse();
        } else if (lineOrder === "uniqueAdjacent") {
            var uniqueLines = [];
            for (var i = 0; i < lines.length; i++) {
                if (i === 0 || lines[i] !== lines[i - 1]) {
                    uniqueLines.push(lines[i]);
                }
            }
            lines = uniqueLines;
        }
        return lines.join("\r");
    }

    /**
     * 変換に使う正規表現をまとめて作る（行頭・行末スペースは［連続するスペース］の設定で変わる）
     * @param {boolean} collapseRepeatedSpaces - 連続するスペースもまとめて削除するなら true
     * @returns {Object} 正規表現の一式
     */
    function buildRuntimeRegexes(collapseRepeatedSpaces) {
        /* 半角スペース + 全角スペース（常に対象） / Half- and full-width spaces */
        var spaceChars = " \u3000";
        var quantifier = collapseRepeatedSpaces ? "+" : "";
        var digits = "0-9\\uFF10-\\uFF19"; /* 半角/全角数字 / half- and full-width digits */

        return {
            reLead: new RegExp("(^|[\\r\\n\\u0003])[" + spaceChars + "]" + quantifier, "g"),
            reTrail: new RegExp("[" + spaceChars + "]" + quantifier + "([\\r\\n\\u0003]|$)", "g"),
            /* 行頭マーカー: 1. / 1 / A. / ① / ・ / - など（行頭 or 改行直後）と直後の空白
               Line-head markers such as 1. / 1 / A. / ① / ・ / - and the blanks after them */
            reHeadMarker: new RegExp(
                "(^|[\\r\\n\\u0003])" +
                "(?:" +
                "[" + digits + "]+(?:[\\.\\uFF0E])?" +           /* 1 / 1. / １． */
                "|[A-Za-z](?:[\\.\\uFF0E])" +                   /* A. / b. */
                "|[\\u2460-\\u2473]" +                           /* ①-⑳ */
                "|・" +                                            /* ・ */
                "|-" +                                             /* - */
                ")" +
                "[" + spaceChars + "\\t]*",                       /* 直後の空白（半角/全角/タブ） */
                "g"
            ),
            /* 段落改行は "\r"、強制改行（Shift+Return）は "\u0003"。"\n" が混じることもあるので強制改行として扱う
               Paragraph breaks are \r and forced breaks are \u0003; a stray \n counts as a forced break */
            FORCED_BR: "\u0003",
            reForced: /[\n\u0003]/g,
            rePara: /\r/g,
            /* 2つ以上の段落改行（空白だけの行を含む）を1つに / Two or more paragraph breaks, blank-only lines included */
            reBlankPara2Plus: /\r(?:[ \u3000\t]*\r)+/g,
            reTab: /\t/g,
            reHyphen: /-/g,
            reUnderscore: /_/g
        };
    }

    /**
     * 置換の第1グループ（行頭や改行）だけを残す
     * @param {string} _match - 一致した文字列
     * @param {string} keptPrefix - 第1グループ
     * @returns {string} 第1グループ（無ければ空文字）
     */
    function keepFirstGroup(_match, keptPrefix) {
        return (keptPrefix != null ? keptPrefix : "");
    }

    /**
     * 設定に従って文字列を整形する
     * @param {string} sourceText - 対象の文字列
     * @param {Object} normalizeOptions - readNormalizeOptions() の結果
     * @returns {string} 整形した文字列
     */
    function normalizeText(sourceText, normalizeOptions) {
        var converted = sourceText;
        var patterns = buildRuntimeRegexes(normalizeOptions.collapseSpaces);

        /* 1) スペースの行頭・行末削除 / Trim leading and trailing spaces */
        if (normalizeOptions.trimLeading) converted = converted.replace(patterns.reLead, keepFirstGroup);
        if (normalizeOptions.trimTrailing) converted = converted.replace(patterns.reTrail, keepFirstGroup);

        /* 2) 行頭マーカー削除（番号/記号） / Remove line-head markers */
        if (normalizeOptions.removeMarkers) converted = converted.replace(patterns.reHeadMarker, keepFirstGroup);

        /* 3) 改行変換（順序注意） / Convert breaks (order matters) */
        if (normalizeOptions.forcedToPara) {
            converted = converted.replace(patterns.reForced, "\r");
        } else if (normalizeOptions.paraToForced) {
            converted = converted.replace(patterns.rePara, patterns.FORCED_BR);
        }

        /* 4) 空行の整理（連続改行を \r 1つに圧縮） / Collapse blank lines into a single \r */
        if (normalizeOptions.compressBlank) converted = converted.replace(patterns.reBlankPara2Plus, "\r");

        /* 5) タブ→半角スペース / Tabs to spaces */
        if (normalizeOptions.tabToSpace) converted = converted.replace(patterns.reTab, " ");

        /* 5.4) ハイフン／アンダースコア変換（両方オンなら何もしない） / Hyphen <-> underscore (neither when both are on) */
        if (normalizeOptions.hyphenToUnderscore && !normalizeOptions.underscoreToHyphen) {
            converted = converted.replace(patterns.reHyphen, "_");
        } else if (normalizeOptions.underscoreToHyphen && !normalizeOptions.hyphenToUnderscore) {
            converted = converted.replace(patterns.reUnderscore, "-");
        }

        /* 5.5) 全角英数字を半角に / Full-width alphanumerics to half-width */
        if (normalizeOptions.zenkakuToHankaku) converted = zenkakuAlnumToHankaku(converted);

        /* 5.6) 英字のケース変換 / Letter case */
        converted = convertCase(converted, normalizeOptions.caseMode);

        /* 5.7) ソート／逆順／重複削除 / Sort, reverse or remove adjacent duplicates */
        if (normalizeOptions.lineOrder) converted = reorderLines(converted, normalizeOptions.lineOrder);

        /* 6) ナンバリングの振り直し。削除と同時指定なら削除を優先 / Renumber; removal wins when both are set */
        if (normalizeOptions.renumber && !normalizeOptions.removeMarkers) {
            converted = renumberLineHeads(converted, normalizeOptions.renumberStyle);
        }

        return converted;
    }

    /**
     * 変換例の表示用に、改行と連続スペースをまとめて短くする
     * @param {string} exampleText - 元の文字列
     * @returns {string} 表示用の文字列
     */
    function normalizeExampleText(exampleText) {
        if (exampleText == null) return "";
        /* 改行/強制改行をスペースに、連続スペースを軽く整理 / Breaks to spaces, then tidy up spaces */
        var shortened = String(exampleText).replace(/[\r\n\u0003]+/g, " ");
        shortened = shortened.replace(/[ \u3000\t]+/g, " ").replace(/^\s+|\s+$/g, "");
        /* 長すぎるとUIが伸びるので短縮 / Shorten so the dialog does not stretch */
        if (shortened.length > CASE_EXAMPLE_MAX_LENGTH) shortened = shortened.substring(0, CASE_EXAMPLE_MAX_LENGTH) + "…";
        return shortened;
    }

    // =========================================
    // 対象テキスト / Target text
    // =========================================

    /**
     * テキストフレームかどうかを返す
     * @param {Object} targetItem - 判定するオブジェクト
     * @returns {boolean} テキストフレームなら true
     */
    function isTextFrame(targetItem) { return targetItem && targetItem.typename === "TextFrame"; }

    /**
     * テキスト範囲（文字選択）かどうかを返す
     * @param {Object} targetItem - 判定するオブジェクト
     * @returns {boolean} TextRange なら true
     */
    function isTextRange(targetItem) { return targetItem && targetItem.typename === "TextRange"; }

    /**
     * オブジェクト（グループの中身も含む）から TextRange を重複なく集める
     * @param {Object} targetItem - 選択中のオブジェクト
     * @param {TextRange[]} textRanges - 集めた TextRange（ここに追加する）
     * @returns {void}
     */
    function collectTextRanges(targetItem, textRanges) {
        if (!targetItem) return;
        var textRange = null;
        if (isTextRange(targetItem)) {
            textRange = targetItem;
        } else if (isTextFrame(targetItem)) {
            textRange = targetItem.textRange;
        } else {
            if (targetItem.typename === "GroupItem") {
                for (var i = 0; i < targetItem.pageItems.length; i++) collectTextRanges(targetItem.pageItems[i], textRanges);
            }
            return;
        }
        for (var j = 0; j < textRanges.length; j++) {
            if (textRanges[j] === textRange) return;
        }
        textRanges.push(textRange);
    }

    /**
     * 対象のすべての TextRange を集める
     * @param {Object[]} targetItems - 対象のオブジェクト
     * @returns {TextRange[]} TextRange の一覧
     */
    function collectAllTextRanges(targetItems) {
        var textRanges = [];
        for (var i = 0; i < targetItems.length; i++) collectTextRanges(targetItems[i], textRanges);
        return textRanges;
    }

    /**
     * 処理の対象を返す。選択が無ければドキュメント内のすべてのテキストフレーム
     * @returns {Object[]} 対象のオブジェクト
     */
    function getTargetItems() {
        var targetItems = app.selection;
        if (!targetItems || targetItems.length === 0) {
            targetItems = [];
            var allFrames = app.activeDocument.textFrames;
            for (var i = 0; i < allFrames.length; i++) {
                targetItems.push(allFrames[i]);
            }
        }
        return targetItems;
    }

    /**
     * 対象の現在の文字列を控える（リセットで戻すため）
     * @param {Object[]} targetItems - 対象のオブジェクト
     * @returns {Object[]} { textRange, baseline } の一覧
     */
    function takeBaseline(targetItems) {
        var textRanges = collectAllTextRanges(targetItems);
        var baselineRanges = [];
        for (var i = 0; i < textRanges.length; i++) {
            /* 読めない範囲は控えない / skip ranges that cannot be read */
            try { baselineRanges.push({ textRange: textRanges[i], baseline: textRanges[i].contents }); } catch (e) { }
        }
        return baselineRanges;
    }

    /**
     * 控えた文字列を書き戻す
     * @param {Object[]} baselineRanges - takeBaseline() の結果
     * @returns {void}
     */
    function restoreBaseline(baselineRanges) {
        for (var i = 0; i < baselineRanges.length; i++) {
            /* 削除済みなどで書けない範囲は飛ばす / skip ranges that can no longer be written */
            try { baselineRanges[i].textRange.contents = baselineRanges[i].baseline; } catch (e) { }
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * タブ（またはパネル）を縦並び・左揃えに整える
     * @param {Object} container - 対象の tab / panel
     * @returns {Object} 同じコンテナ
     */
    function setupColumnContainer(container) {
        container.orientation = "column";
        container.alignChildren = ["left", "top"];
        container.margins = PANEL_MARGINS;
        return container;
    }

    /**
     * 縦並び・左揃えのグループを追加する
     * @param {Object} parentContainer - 追加先
     * @returns {Group} 追加したグループ
     */
    function addColumnGroup(parentContainer) {
        var columnGroup = parentContainer.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["left", "top"];
        return columnGroup;
    }

    /**
     * 少し小さめのボタンを追加する
     * @param {Object} parentContainer - 追加先
     * @param {string} labelPath - ラベルのパス
     * @param {string} tooltipPath - ツールチップのラベルパス
     * @returns {Button} 追加したボタン
     */
    function addSmallButton(parentContainer, labelPath, tooltipPath) {
        var smallButton = parentContainer.add("button", undefined, getLabel(labelPath));
        smallButton.helpTip = getLabel(tooltipPath);
        smallButton.preferredSize.height = SMALL_BUTTON_HEIGHT;
        return smallButton;
    }

    /**
     * ツールチップ付きのチェックボックスを追加する
     * @param {Object} parentContainer - 追加先
     * @param {string} labelPath - ラベルのパス
     * @param {string} tooltipPath - ツールチップのラベルパス
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parentContainer, labelPath, tooltipPath, initialValue) {
        var checkbox = parentContainer.add("checkbox", undefined, getLabel(labelPath));
        checkbox.value = initialValue;
        checkbox.helpTip = getLabel(tooltipPath);
        return checkbox;
    }

    /* 英字のケース変換ボタンの並び（mode とラベル） / Case buttons: mode and label */
    var CASE_BUTTONS = [
        { mode: "upper", labelPath: "button.caseUpper" },
        { mode: "lower", labelPath: "button.caseLower" },
        { mode: "word", labelPath: "button.caseWord" },
        { mode: "sentence", labelPath: "button.caseSentence" },
        { mode: "title", labelPath: "button.caseTitle" }
    ];

    /* 振り直しの形式ラジオ（style とラベル） / Numbering style radios: style and label */
    var NUMBER_STYLE_RADIOS = [
        { style: "num", labelPath: "radio.styleNumDot" },
        { style: "alpha", labelPath: "radio.styleAlphaDot" },
        { style: "circled", labelPath: "radio.styleCircled" },
        { style: "dot", labelPath: "radio.styleDot" },
        { style: "hyphen", labelPath: "radio.styleHyphen" }
    ];

    /**
     * ダイアログとコントロールを組み立てる（イベントは bindDialogEvents() で付ける）
     * @returns {Object} ダイアログ本体（window）と各コントロール
     */
    function buildDialog() {
        var ui = {};
        ui.window = new Window("dialog", getLabel("dialog.title"));
        ui.window.orientation = "column";
        ui.window.alignChildren = ["fill", "top"];
        ui.window.preferredSize.width = DIALOG_WIDTH;

        var settingsTabs = ui.window.add("tabbedpanel");
        settingsTabs.alignChildren = ["fill", "top"];
        settingsTabs.preferredSize.width = DIALOG_WIDTH;

        /* 削除タブ：行頭/行末、連続スペース / Remove tab: leading/trailing and repeated spaces */
        var removeTab = setupColumnContainer(settingsTabs.add("tab", undefined, getLabel("tab.remove")));
        ui.cbLeadingSpace = addCheckbox(removeTab, "checkbox.leadingSpace", "tooltip.leadingSpace", true);
        ui.cbTrailingSpace = addCheckbox(removeTab, "checkbox.trailingSpace", "tooltip.trailingSpace", true);
        ui.cbMultipleSpace = addCheckbox(removeTab, "checkbox.multipleSpace", "tooltip.multipleSpace", true);

        /* ナンバリングタブ：2カラム（左：ボタン / 右：形式） / Numbering tab: buttons on the left, format on the right */
        var numberingTab = setupColumnContainer(settingsTabs.add("tab", undefined, getLabel("tab.numbering")));
        var numberingColumns = numberingTab.add("group");
        numberingColumns.orientation = "row";
        numberingColumns.alignChildren = ["fill", "top"];
        var numberingButtonColumn = addColumnGroup(numberingColumns);
        var numberingStyleColumn = addColumnGroup(numberingColumns);

        var numberingButtonGroup = addColumnGroup(numberingButtonColumn);
        ui.btnRemoveMarker = addSmallButton(numberingButtonGroup, "button.removeMarker", "tooltip.removeMarker");
        ui.btnRenumber = addSmallButton(numberingButtonGroup, "button.renumber", "tooltip.renumber");
        ui.btnResetNumbering = addSmallButton(numberingButtonGroup, "button.reset", "tooltip.resetNumbering");

        var numberStylePanel = setupColumnContainer(numberingStyleColumn.add("panel", undefined, getLabel("panel.numberStyle")));
        ui.numberStyleRadios = [];
        for (var i = 0; i < NUMBER_STYLE_RADIOS.length; i++) {
            ui.numberStyleRadios.push(numberStylePanel.add("radiobutton", undefined, getLabel(NUMBER_STYLE_RADIOS[i].labelPath)));
        }
        ui.rbNumDot = ui.numberStyleRadios[0];
        /* デフォルト: 1. 形式 / Default: 1. */
        ui.rbNumDot.value = true;
        for (var r = 0; r < ui.numberStyleRadios.length; r++) {
            ui.numberStyleRadios[r].helpTip = getLabel("tooltip.numberStyle");
        }

        /* 行タブ：改行変換、空行整理、ソート / Lines tab: break conversion, blank lines, sort */
        var linesTab = setupColumnContainer(settingsTabs.add("tab", undefined, getLabel("tab.lines")));
        var lineBreakButtonGroup = addColumnGroup(linesTab);
        ui.btnForcedToPara = addSmallButton(lineBreakButtonGroup, "button.forcedToPara", "tooltip.forcedToPara");
        ui.btnParaToForced = addSmallButton(lineBreakButtonGroup, "button.paraToForced", "tooltip.paraToForced");
        ui.btnCompressBlank = addSmallButton(linesTab, "button.compressBlank", "tooltip.compressBlank");

        var sortPanel = setupColumnContainer(linesTab.add("panel", undefined, getLabel("panel.sort")));
        var sortGroup = addColumnGroup(sortPanel);
        ui.rbSortAsc = sortGroup.add("radiobutton", undefined, getLabel("radio.sortAsc"));
        ui.rbReverse = sortGroup.add("radiobutton", undefined, getLabel("radio.reverse"));
        ui.rbUniqueAdjacent = sortGroup.add("radiobutton", undefined, getLabel("radio.uniqueAdjacent"));
        ui.rbSortAsc.helpTip = getLabel("tooltip.sortAsc");
        ui.rbReverse.helpTip = getLabel("tooltip.reverse");
        ui.rbUniqueAdjacent.helpTip = getLabel("tooltip.uniqueAdjacent");
        ui.rbSortAsc.value = false;
        ui.rbReverse.value = false;
        ui.rbUniqueAdjacent.value = false;

        /* アルファベットタブ：ケース変換ボタンと変換例 / Letter case tab: case buttons with examples */
        var caseTab = setupColumnContainer(settingsTabs.add("tab", undefined, getLabel("tab.letterCase")));
        var caseGroup = addColumnGroup(caseTab);
        ui.caseButtons = [];
        ui.caseExampleLabels = {};
        for (var c = 0; c < CASE_BUTTONS.length; c++) {
            var caseRow = caseGroup.add("group");
            caseRow.orientation = "row";
            caseRow.alignChildren = ["left", "center"];

            var caseButton = caseRow.add("button", undefined, getLabel(CASE_BUTTONS[c].labelPath));
            caseButton.helpTip = getLabel("tooltip.caseButton");
            caseButton.preferredSize.height = SMALL_BUTTON_HEIGHT;
            caseButton.preferredSize.width = CASE_BUTTON_WIDTH;
            ui.caseButtons.push(caseButton);

            var caseExample = caseRow.add("statictext", undefined, "");
            caseExample.preferredSize.width = CASE_EXAMPLE_WIDTH;
            caseExample.justify = "left";
            ui.caseExampleLabels[CASE_BUTTONS[c].mode] = caseExample;
        }
        ui.btnResetCase = addSmallButton(caseGroup, "button.reset", "tooltip.resetCase");

        /* その他タブ / Other tab */
        var otherTab = setupColumnContainer(settingsTabs.add("tab", undefined, getLabel("tab.other")));
        ui.cbTabToSpace = addCheckbox(otherTab, "checkbox.tabToSpace", "tooltip.tabToSpace", false);
        ui.cbZenkakuToHankaku = addCheckbox(otherTab, "checkbox.zenkakuToHankaku", "tooltip.zenkakuToHankaku", false);
        ui.cbHyphenToUnderscore = addCheckbox(otherTab, "checkbox.hyphenToUnderscore", "tooltip.hyphenToUnderscore", false);
        ui.cbUnderscoreToHyphen = addCheckbox(otherTab, "checkbox.underscoreToHyphen", "tooltip.underscoreToHyphen", false);

        /* フッター（プレビュー＋ボタン） / Footer (preview and buttons) */
        var footerGroup = ui.window.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignChildren = ["fill", "center"];

        var footerLeftGroup = footerGroup.add("group");
        footerLeftGroup.orientation = "row";
        footerLeftGroup.alignChildren = ["left", "center"];
        footerLeftGroup.alignment = ["left", "center"];
        ui.cbPreview = addCheckbox(footerLeftGroup, "checkbox.preview", "tooltip.preview", true);

        var btnRightGroup = footerGroup.add("group");
        btnRightGroup.orientation = "row";
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.alignment = ["right", "center"];
        ui.btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        ui.btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return ui;
    }

    /**
     * ダイアログの状態から整形の設定を読み取る
     * @param {Object} ui - buildDialog() の結果
     * @param {Object} normalizeState - ボタンで切り替える設定（行頭マーカー・改行・ケースなど）
     * @returns {Object} normalizeText() に渡す設定
     */
    function readNormalizeOptions(ui, normalizeState) {
        var lineOrder = null;
        if (ui.rbSortAsc.value) lineOrder = "sortAsc";
        else if (ui.rbReverse.value) lineOrder = "reverse";
        else if (ui.rbUniqueAdjacent.value) lineOrder = "uniqueAdjacent";

        return {
            trimLeading: ui.cbLeadingSpace.value,
            trimTrailing: ui.cbTrailingSpace.value,
            collapseSpaces: ui.cbMultipleSpace.value,
            removeMarkers: normalizeState.removeMarkers,
            forcedToPara: normalizeState.forcedToPara,
            paraToForced: normalizeState.paraToForced,
            compressBlank: normalizeState.compressBlank,
            tabToSpace: ui.cbTabToSpace.value,
            hyphenToUnderscore: ui.cbHyphenToUnderscore.value,
            underscoreToHyphen: ui.cbUnderscoreToHyphen.value,
            zenkakuToHankaku: ui.cbZenkakuToHankaku.value,
            caseMode: normalizeState.caseMode,
            lineOrder: lineOrder,
            renumber: normalizeState.renumber,
            renumberStyle: normalizeState.renumberStyle
        };
    }

    /**
     * 実行する処理が1つでも選ばれているかを返す（［連続するスペース］単独は数えない）
     * @param {Object} normalizeOptions - readNormalizeOptions() の結果
     * @returns {boolean} 選ばれていれば true
     */
    function isAnyOptionSelected(normalizeOptions) {
        return !!(
            normalizeOptions.trimLeading || normalizeOptions.trimTrailing || normalizeOptions.removeMarkers ||
            normalizeOptions.forcedToPara || normalizeOptions.paraToForced ||
            normalizeOptions.compressBlank || normalizeOptions.tabToSpace || normalizeOptions.renumber ||
            normalizeOptions.zenkakuToHankaku ||
            (normalizeOptions.caseMode != null) ||
            normalizeOptions.hyphenToUnderscore || normalizeOptions.underscoreToHyphen ||
            normalizeOptions.lineOrder
        );
    }

    /**
     * ダイアログの各コントロールにイベントを付ける（プレビュー・リセット・OK）
     * @param {Object} ui - buildDialog() の結果
     * @param {Object[]} targetItems - 処理の対象
     * @returns {void}
     */
    function bindDialogEvents(ui, targetItems) {
        /* ボタンで切り替える設定 / Settings toggled by buttons */
        var normalizeState = {
            removeMarkers: false,  /* 行頭マーカー削除 / remove line markers */
            renumber: false,       /* ナンバリングの振り直し / renumber */
            renumberStyle: "num",  /* "num"=1. / "alpha"=A. / "circled"=① / "dot"=・ / "hyphen"=- */
            forcedToPara: false,   /* 強制改行→改行 / forced breaks to paragraph breaks */
            paraToForced: false,   /* 改行→強制改行 / paragraph breaks to forced breaks */
            compressBlank: false,  /* 空行整理 / collapse blank lines */
            caseMode: null         /* "upper" | "lower" | "word" | "sentence" | "title"（null=未選択） */
        };
        /* ダイアログ表示時点（初回プレビュー適用後）の状態 / State when the dialog opened (after the first preview) */
        var baselineRanges = [];
        var baselineReady = false;

        /**
         * 対象のすべてのテキストを整形する
         * @returns {void}
         */
        function applyProcessToSelection() {
            var normalizeOptions = readNormalizeOptions(ui, normalizeState);
            var textRanges = collectAllTextRanges(targetItems);
            for (var j = 0; j < textRanges.length; j++) {
                textRanges[j].contents = normalizeText(textRanges[j].contents, normalizeOptions);
            }
        }

        /**
         * プレビューがオンなら整形して再描画する（簡易プレビュー：ヒストリーは積み重なる）
         * @returns {void}
         */
        function requestPreview() {
            if (!ui.cbPreview.value) return;
            /* 書き込めないテキストがあってもダイアログは閉じない / keep the dialog open even if a write fails */
            try {
                applyProcessToSelection();
                app.redraw(); /* プレビューON時のみ強制再描画 / redraw only while previewing */
            } catch (e) { }
        }

        /**
         * 変換例の元になる文字列を返す（控えがあれば最初の控え、無ければ現在の最初のテキスト）
         * @returns {string} 元の文字列
         */
        function getFirstSelectedTextSnapshot() {
            if (baselineReady && baselineRanges.length > 0) return baselineRanges[0].baseline;
            /* 対象が読めなければ空にする / fall back to empty when the text cannot be read */
            try {
                var textRanges = collectAllTextRanges(targetItems);
                if (textRanges.length > 0) return textRanges[0].contents;
            } catch (e) { }
            return "";
        }

        /**
         * ケース変換ボタンの右に変換例を表示する
         * @returns {void}
         */
        function updateCaseExamples() {
            var exampleBaseText = normalizeExampleText(getFirstSelectedTextSnapshot());
            for (var c = 0; c < CASE_BUTTONS.length; c++) {
                var exampleLabel = ui.caseExampleLabels[CASE_BUTTONS[c].mode];
                if (exampleLabel) exampleLabel.text = normalizeExampleText(convertCase(exampleBaseText, CASE_BUTTONS[c].mode));
            }
        }

        /**
         * 押したときに設定を切り替えてプレビューするボタンにする
         * @param {Button} stateButton - 対象のボタン
         * @param {Function} updateState - 設定を切り替える処理
         * @returns {void}
         */
        function bindStateButton(stateButton, updateState) {
            stateButton.onClick = function () {
                updateState();
                requestPreview();
            };
        }

        /* ナンバリング / Numbering */
        bindStateButton(ui.btnRemoveMarker, function () {
            normalizeState.removeMarkers = true;
            normalizeState.renumber = false;
        });
        bindStateButton(ui.btnRenumber, function () {
            normalizeState.renumber = true;
            normalizeState.removeMarkers = false;
        });
        ui.btnResetNumbering.onClick = function () {
            normalizeState.removeMarkers = false;
            normalizeState.renumber = false;
            normalizeState.renumberStyle = "num";
            ui.rbNumDot.value = true;

            /* プレビューON時：ナンバリング処理で失われた行頭マーカー等を戻すため、ベースラインへ復元してから再適用する
               While previewing, restore the baseline first so markers lost to numbering come back, then reapply */
            if (ui.cbPreview && ui.cbPreview.value) {
                restoreBaseline(baselineRanges);
            }

            requestPreview();
            updateCaseExamples();
        };
        for (var r = 0; r < ui.numberStyleRadios.length; r++) {
            (function (renumberStyle) {
                ui.numberStyleRadios[r].onClick = function () { normalizeState.renumberStyle = renumberStyle; };
            })(NUMBER_STYLE_RADIOS[r].style);
        }

        /* 行 / Lines */
        bindStateButton(ui.btnForcedToPara, function () {
            normalizeState.forcedToPara = true;
            normalizeState.paraToForced = false;
        });
        bindStateButton(ui.btnParaToForced, function () {
            normalizeState.paraToForced = true;
            normalizeState.forcedToPara = false;
        });
        bindStateButton(ui.btnCompressBlank, function () {
            normalizeState.compressBlank = true;
        });

        /* 英字のケース変換 / Letter case */
        for (var c = 0; c < ui.caseButtons.length; c++) {
            (function (caseMode) {
                bindStateButton(ui.caseButtons[c], function () { normalizeState.caseMode = caseMode; });
            })(CASE_BUTTONS[c].mode);
        }
        ui.btnResetCase.onClick = function () {
            /* テキストを「ダイアログを開いた時点」に戻す / Put the text back as it was when the dialog opened */
            restoreBaseline(baselineRanges);

            /* UI を初期値へ / Reset the UI */
            ui.cbLeadingSpace.value = true;
            ui.cbTrailingSpace.value = true;
            ui.cbMultipleSpace.value = true;

            normalizeState.removeMarkers = false;
            normalizeState.renumber = false;
            normalizeState.renumberStyle = "num";
            ui.rbNumDot.value = true;

            normalizeState.forcedToPara = false;
            normalizeState.paraToForced = false;
            normalizeState.compressBlank = false;

            ui.cbTabToSpace.value = false;
            ui.cbZenkakuToHankaku.value = false;
            ui.cbHyphenToUnderscore.value = false;
            ui.cbUnderscoreToHyphen.value = false;

            ui.rbSortAsc.value = false;
            ui.rbReverse.value = false;
            ui.rbUniqueAdjacent.value = false;

            normalizeState.caseMode = null;

            ui.cbPreview.value = true;

            /* 初期状態（=プレビューON前提）として再描画のみ / Only redraw; the defaults assume preview is on */
            app.redraw();
            updateCaseExamples();
        };

        /* UI 変更時にプレビュー更新 / Update the preview when a setting changes */
        var previewControls = [ui.cbLeadingSpace, ui.cbTrailingSpace, ui.cbMultipleSpace,
            ui.cbTabToSpace, ui.cbZenkakuToHankaku, ui.cbHyphenToUnderscore, ui.cbUnderscoreToHyphen,
            ui.rbSortAsc, ui.rbReverse, ui.rbUniqueAdjacent];
        for (var p = 0; p < previewControls.length; p++) {
            previewControls[p].onClick = requestPreview;
        }
        ui.cbPreview.onClick = function () {
            /* OFF にしても復元はしない（簡易プレビュー） / Turning it off does not restore (simple preview) */
            if (ui.cbPreview.value) requestPreview();
        };

        ui.btnOK.onClick = function () {
            if (!isAnyOptionSelected(readNormalizeOptions(ui, normalizeState))) {
                alert(getLabel("alert.noOption"));
                return;
            }
            /* プレビューONなら適用済み。OFFなら、OK時に1回だけ適用 / Already applied while previewing; otherwise apply once now */
            if (!ui.cbPreview.value) {
                /* 書き込めないテキストがあってもダイアログは閉じる / close the dialog even if a write fails */
                try { applyProcessToSelection(); } catch (e) { }
            }
            ui.window.close(1);
        };
        ui.btnCancel.onClick = function () {
            ui.window.close(0);
        };

        /* 初回表示時：プレビューがONならまず適用し、その結果を「ベースライン」として記録
           On first show: apply the preview if it is on, then record the result as the baseline */
        ui.window.onShow = function () {
            if (ui.cbPreview.value) {
                requestPreview();
            }
            baselineRanges = takeBaseline(targetItems);
            baselineReady = true;
            updateCaseExamples();
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを表示し、選択中（無ければドキュメント内すべて）のテキストを整形する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var targetItems = getTargetItems();

        var ui = buildDialog();
        bindDialogEvents(ui, targetItems);
        /* 整形はプレビューと OK ボタンの中で済んでいる / The work is done in the preview and the OK handler */
        ui.window.show();
    }

    main();
})();
