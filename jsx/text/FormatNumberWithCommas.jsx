#target illustrator
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

テキスト内の数値に桁区切りのカンマを付け、位置の誤ったカンマを直します。
西暦・郵便番号・電話番号などは除外し、書き換える数値はプレビューで選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FormatNumberWithCommas.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n21f07978f177

### Overview

Adds thousands separators to numbers in text and fixes misplaced commas.
Years, postal codes, phone numbers and the like are skipped, and you pick the numbers to change in a preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FormatNumberWithCommas.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FormatNumberWithCommas";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-12";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FormatNumberWithCommas.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FormatNumberWithCommas.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n21f07978f177"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 除外ルール。この順にチェックボックスを並べ、defaultValue を初期状態にする
       Exclusion rules in checkbox order; defaultValue is the initial state */
    var EXCLUSION_RULES = [
        { key: "years", defaultValue: true },          /* 西暦 / years */
        { key: "postalCodes", defaultValue: true },    /* 郵便番号 / postal codes */
        { key: "slashAdjacent", defaultValue: true },  /* スラッシュの前後 / next to a slash */
        { key: "phoneNumbers", defaultValue: true },   /* 電話番号 / phone numbers */
        { key: "creditCards", defaultValue: true },    /* クレジットカード番号 / credit card numbers */
        { key: "macAddresses", defaultValue: true }    /* MACアドレス / MAC addresses */
    ];

    /* プレビューの # を上から順に振る（false で下から）/ Number texts from the top in the preview (false: from the bottom) */
    var RANK_FROM_TOP = true;

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS = [15, 20, 15, 10];       /* パネル余白 [左,上,右,下] / panel margins */
    var DIALOG_OFFSET_X = 300;                  /* ダイアログを右へずらす量 / horizontal dialog offset */
    var DIALOG_OFFSET_Y = 0;                    /* ダイアログを下へずらす量 / vertical dialog offset */
    var DIALOG_OPACITY = 0.98;                  /* ダイアログの不透明度 / dialog opacity */
    var PREVIEW_LIST_WIDTH = 260;               /* プレビューリストの幅 / preview list width */
    var PREVIEW_COLUMN_WIDTHS = [40, 70, 150];  /* 列幅 [選択, #, 数値] / column widths */
    var PREVIEW_HEADER_HEIGHT = 24;             /* 見出し行の高さ / header row height */
    var PREVIEW_ROW_HEIGHT = 20;                /* 1行の高さ / row height */
    var PREVIEW_LIST_PADDING = 40;              /* リストの上下の余白 / list padding */
    var PREVIEW_LIST_MAX_HEIGHT = 520;          /* リストの最大の高さ / max list height */
    var PREVIEW_MIN_ROWS = 4;                   /* 最低限見せる行数 / minimum visible rows */
    var PREVIEW_MAX_ROWS = 16;                  /* スクロールせずに見せる行数 / rows shown before scrolling */
    var CHECK_MARK = "✓";                       /* カンマを付ける行の印 / mark on rows that get commas */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UIの表示言語を返す
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            scopeTitle: { ja: "桁区切りのカンマを付ける", en: "Add Thousands Separators" },
            previewTitle: { ja: "カンマ付与対象の確認", en: "Review Numbers to Add Commas" }
        },
        panel: {
            scope: { ja: "対象", en: "Scope" },
            exclusion: { ja: "除外", en: "Exclude" }
        },
        radio: {
            selection: { ja: "選択したオブジェクトのみ", en: "Selection only" },
            wholeDocument: { ja: "ドキュメントすべて", en: "Whole document" }
        },
        checkbox: {
            years: { ja: "西暦", en: "Years" },
            postalCodes: { ja: "郵便番号", en: "Postal codes" },
            slashAdjacent: { ja: "スラッシュの前後", en: "Next to a slash" },
            phoneNumbers: { ja: "電話番号", en: "Phone numbers" },
            creditCards: { ja: "クレジットカード番号", en: "Credit card numbers" },
            macAddresses: { ja: "MACアドレス", en: "MAC addresses" }
        },
        message: {
            previewInstruction: { ja: "カンマを付ける（直す）数値を選択してください", en: "Select numbers to add/fix commas" }
        },
        listColumn: {
            select: { ja: "選択", en: "Select" },
            frameNumber: { ja: "#", en: "#" },
            value: { ja: "数値", en: "Value" }
        },
        tooltip: {
            years: {
                ja: "直後に「年」があるか、前後に日付の区切り（/ . -）がある4桁の数値（1000〜2999）と、「2024.12」「2024.12.31」のような日付の西暦を除外します。",
                en: "Skips 4-digit numbers (1000–2999) followed by 年 or next to a date separator (/ . -), and years in dates such as 2024.12 or 2024.12.31."
            },
            postalCodes: {
                ja: "「123-4567」の形の数値を、〒の有無にかかわらず除外します。",
                en: "Skips numbers in the 123-4567 form, with or without 〒."
            },
            slashAdjacent: {
                ja: "直前か直後にスラッシュ（/ ／）がある数値を除外します。",
                en: "Skips numbers directly before or after a slash (/ ／)."
            },
            phoneNumbers: {
                ja: "0・+・括弧で始まり、ハイフン・スペース・括弧で区切った電話番号を除外します。",
                en: "Skips phone numbers that start with 0, + or a parenthesis and are split by hyphens, spaces or parentheses."
            },
            creditCards: {
                ja: "ハイフンやスペースで区切ったものを含め、13〜19桁でカード番号の検査（Luhn チェック）を通る数値を除外します。",
                en: "Skips 13–19 digit numbers, including ones split by hyphens or spaces, that pass the Luhn check used for card numbers."
            },
            macAddresses: {
                ja: "「00:1A:2B:3C:4D:5E」のようにコロンかハイフンで区切ったMACアドレスを除外します。",
                en: "Skips MAC addresses separated by colons or hyphens, such as 00:1A:2B:3C:4D:5E."
            },
            previewList: {
                ja: "✓ の付いた行の数値にだけカンマを付けます。行をクリックすると ✓ を付け外しします。# はテキストの番号（上から順）です。",
                en: "Only rows with ✓ get commas. Click a row to toggle its ✓. # numbers each text from the top."
            },
            selectAll: { ja: "すべての行に ✓ を付けます。", en: "Puts ✓ on every row." }
        },
        button: {
            selectAll: { ja: "すべて選択", en: "Select All" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noTextFrames: { ja: "対象のテキストが見つかりませんでした。", en: "No text objects were found in the target." },
            noCandidates: { ja: "カンマを付ける（直す）数値は見つかりませんでした。", en: "No numbers need commas added or fixed." }
        }
    };

    /**
     * ラベル定義から現在の言語の文言を取り出す
     * @param {object} labelSet - { ja: string, en: string } 形式のラベル定義
     * @returns {string} 現在の言語の文言
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || "";
    }

    // =========================================
    // 数値の整形 / Number formatting
    // =========================================

    /**
     * 半角・全角のカンマか
     * @param {string} charText - 1文字
     * @returns {boolean} カンマなら true
     */
    function isComma(charText) {
        return charText === "," || charText === "，";
    }

    /**
     * 半角・全角のカンマをすべて取り除く
     * @param {string} text - 対象の文字列
     * @returns {string} カンマを除いた文字列
     */
    function removeCommas(text) {
        return text.replace(/[，,]/g, "");
    }

    /**
     * 数字の並びに3桁ごとのカンマを入れる（符号・小数部は含まない）
     * @param {string} digits - 数字だけの文字列
     * @param {string} comma - 入れるカンマ（"," または "，"）
     * @returns {string} カンマを入れた文字列
     */
    function insertCommasToDigits(digits, comma) {
        var groupedDigits = "";
        for (var i = 0; i < digits.length; i++) {
            if (i > 0 && (digits.length - i) % 3 === 0) groupedDigits += comma;
            groupedDigits += digits.charAt(i);
        }
        return groupedDigits;
    }

    /**
     * 数値のトークンを符号・整数部・小数部に分ける
     * @param {string} token - 数値のトークン
     * @returns {{sign: string, integerPart: string, integerDigits: string, fraction: string}}
     *     integerPart はカンマ込み、integerDigits はカンマなし、fraction は小数点込み
     */
    function splitNumberToken(token) {
        var sign = /^[+\-＋－]/.test(token) ? token.charAt(0) : "";
        var unsignedText = token.substring(sign.length);
        var pointIndex = unsignedText.search(/[\.．]/);
        var integerPart = (pointIndex < 0) ? unsignedText : unsignedText.substring(0, pointIndex);
        return {
            sign: sign,
            integerPart: integerPart,
            integerDigits: removeCommas(integerPart),
            fraction: (pointIndex < 0) ? "" : unsignedText.substring(pointIndex)
        };
    }

    /**
     * 整数部が全角数字か（先頭の数字で決める）
     * @param {string} integerDigits - カンマを除いた整数部
     * @returns {boolean} 全角なら true
     */
    function isFullWidthDigits(integerDigits) {
        return /^[０-９]/.test(integerDigits);
    }

    /**
     * 符号・小数点を、半角か全角の一方にそろえる
     * @param {string} text - 符号または小数部
     * @param {boolean} fullWidth - 全角にそろえるなら true
     * @returns {string} そろえた文字列
     */
    function matchSymbolWidth(text, fullWidth) {
        return fullWidth ? text.replace(/[+\-.]/g, toFullWidthChar) : text.replace(/[＋－．]/g, toHalfWidthChar);
    }

    /**
     * 数値を桁区切りの形にする。カンマ・符号・小数点は数字の幅（半角／全角）にそろえる
     * @param {string} token - 元のテキストにあるままの数値
     * @returns {string} カンマを入れた文字列
     */
    function formatNumberWithCommas(token) {
        var numberParts = splitNumberToken(token);
        var fullWidth = isFullWidthDigits(numberParts.integerDigits);
        return matchSymbolWidth(numberParts.sign, fullWidth) +
            insertCommasToDigits(numberParts.integerDigits, fullWidth ? "，" : ",") +
            matchSymbolWidth(numberParts.fraction, fullWidth);
    }

    /**
     * カンマを付ける、または直す必要があるかを判定する（整数部が4桁以上のときだけ）
     * @param {string} token - 元のテキストにあるままの数値
     * @returns {boolean} 整数部のカンマが正しい桁区切りと異なれば true
     */
    function needsCommaFix(token) {
        var numberParts = splitNumberToken(token);
        if (numberParts.integerDigits.length < 4) return false;
        var comma = isFullWidthDigits(numberParts.integerDigits) ? "，" : ",";
        return insertCommasToDigits(numberParts.integerDigits, comma) !== numberParts.integerPart;
    }

    // =========================================
    // 全角・半角の変換 / Full-width and half-width conversion
    // =========================================

    /**
     * 全角の英数字・記号1文字を半角にする
     * @param {string} fullWidthChar - 全角の1文字
     * @returns {string} 半角の1文字
     */
    function toHalfWidthChar(fullWidthChar) {
        return String.fromCharCode(fullWidthChar.charCodeAt(0) - 0xFEE0);
    }

    /**
     * 半角の英数字・記号1文字を全角にする
     * @param {string} halfWidthChar - 半角の1文字
     * @returns {string} 全角の1文字
     */
    function toFullWidthChar(halfWidthChar) {
        return String.fromCharCode(halfWidthChar.charCodeAt(0) + 0xFEE0);
    }

    /**
     * 数値の全角文字（数字・符号・小数点）を半角にする。除外の判定用
     * @param {string} numberText - カンマを除いた数値
     * @returns {string} 半角にした数値
     */
    function toHalfWidthNumber(numberText) {
        return numberText.replace(/[０-９＋－．]/g, toHalfWidthChar);
    }

    /**
     * 全角の数字とハイフン（－）を半角にする
     * @param {string} text - 対象の文字列
     * @returns {string} 変換後の文字列
     */
    function toHalfWidthDigitsHyphen(text) {
        return text.replace(/[０-９－]/g, toHalfWidthChar);
    }

    /**
     * 電話番号に現れる全角文字（数字・ダッシュ類・括弧・＋・全角スペース）を半角にする
     * @param {string} text - 対象の文字列
     * @returns {string} 変換後の文字列
     */
    function toHalfWidthForPhone(text) {
        return text
            .replace(/[０-９（）＋]/g, toHalfWidthChar)
            .replace(/[－ー―–—]/g, "-")
            .replace(/　/g, " ");
    }

    /**
     * MACアドレスに現れる全角文字（数字・A〜F・コロン・ダッシュ類）を半角にする
     * @param {string} text - 対象の文字列
     * @returns {string} 変換後の文字列
     */
    function toHalfWidthForMac(text) {
        return text
            .replace(/[０-９Ａ-Ｆａ-ｆ：]/g, toHalfWidthChar)
            .replace(/[－ー―–—]/g, "-");
    }

    // =========================================
    // 除外の判定 / Exclusion checks
    // =========================================

    /* 電話番号のパターン（全角は半角にそろえてから照合）/ Phone number patterns, matched after converting to half-width */
    var PHONE_PATTERNS = [
        /^(?:\(\d{2,4}\)|\d{2,4})[-)]?\d{2,4}-\d{4}$/,                         /* (03)1234-5678 */
        /^0[5789]0\d{8}$/,                                                     /* 携帯・IP電話（区切りなし）/ mobile and IP without separators */
        /^0120\d{6}$/,                                                         /* フリーダイヤル（区切りなし）/ toll-free without separators */
        /^\+?\(?\d+\)?(?:[-\s]\(?\d+\)?){1,4}$/,                               /* ハイフン・スペース区切り全般 / groups split by hyphens or spaces */
        /^\(\d+\)\s*\d+(?:[-\s]\d+){1,3}$/,                                    /* (03) 1234 5678 */
        /^\+?\(?\d+\)?(?:[-\s]\(?\d+\)?){1,4}\s*(?:ext\.?|x|内線)\s*\d{1,5}$/  /* 内線付き / with extension */
    ];

    /* 日付の区切り / Date separators */
    var DATE_SEPARATOR = /[\/\.\-－―／．]/;

    /**
     * 数値の前後へ、指定した文字が続く範囲まで広げた文字列を返す
     * @param {string} text - テキスト全体
     * @param {number} start - 数値の開始位置
     * @param {number} end - 数値の終了位置（この位置の文字は含まない）
     * @param {RegExp} charPattern - 1文字を判定する文字クラス
     * @returns {string} 広げた範囲の文字列
     */
    function extractSurroundingToken(text, start, end, charPattern) {
        while (start > 0 && charPattern.test(text.charAt(start - 1))) start--;
        while (end < text.length && charPattern.test(text.charAt(end))) end++;
        return text.substring(start, end);
    }

    /**
     * 空白を飛ばして最初に見つかる文字を返す
     * @param {string} text - 対象の文字列
     * @param {number} index - 探し始める位置
     * @param {number} step - 1 で後ろへ、-1 で前へ進む
     * @returns {string} 見つかった文字。範囲外なら空文字
     */
    function findNonSpaceChar(text, index, step) {
        while (index >= 0 && index < text.length && /\s/.test(text.charAt(index))) index += step;
        return text.charAt(index);
    }

    /**
     * 直前か直後にスラッシュ（/ ／）があるか
     * @param {string} text - テキスト全体
     * @param {number} start - 数値の開始位置
     * @param {number} end - 数値の終了位置
     * @returns {boolean} スラッシュが隣にあれば true
     */
    function isNextToSlash(text, start, end) {
        return /[\/／]/.test(text.charAt(start - 1) + text.charAt(end));
    }

    /**
     * 整数部が0で始まるか（00123456、0312345678、0000 など）。番号やコードとみなす
     * @param {string} value - カンマを除いた半角の数値
     * @returns {boolean} 0で始まれば true
     */
    function hasLeadingZero(value) {
        return /^[-+]?0\d/.test(value);
    }

    /**
     * コードや番号の一部か。英字や # のすぐ後に続く数値（SKU12345、#112233）と、
     * 「数字.」に続く数値（10.0.19045 の 19045）を該当とする。後ろの英字は単位（12000mm など）なので見ない
     * @param {string} text - テキスト全体
     * @param {number} start - 数値の開始位置
     * @returns {boolean} コードや番号の一部なら true
     */
    function isPartOfCode(text, start) {
        var prevChar = text.charAt(start - 1);
        if (/[A-Za-zＡ-Ｚａ-ｚ#＃]/.test(prevChar)) return true;
        return /[\.．]/.test(prevChar) && /[0-9０-９]/.test(text.charAt(start - 2));
    }

    /**
     * 西暦らしいか。1000〜2999 の4桁で、直後に「年」か前後に日付の区切りがあるとき、
     * または「2024.12」「2024.1.5」のような年月のとき西暦とみなす
     * @param {string} value - カンマを除いた半角の数値
     * @param {string} text - テキスト全体
     * @param {number} start - 数値の開始位置
     * @param {number} end - 数値の終了位置
     * @returns {boolean} 西暦とみなせば true
     */
    function looksLikeYear(value, text, start, end) {
        var yearText = value;
        var yearStart = start;
        /* 数字の直後のハイフンは符号ではなく範囲の区切り（2020-2024）/ A hyphen right after a digit separates a range, not a sign */
        if (yearText.charAt(0) === "-" && /[0-9０-９]/.test(text.charAt(start - 1))) {
            yearText = yearText.substring(1);
            yearStart++;
        }

        /* 小数部が2桁の月か、日が続く形だけを年月とみなし、「1500.5」は小数のまま
           Only a 2-digit month, or a month followed by a day, counts as a date; 1500.5 stays a decimal */
        var yearMonth = /^[12]\d{3}\.(\d{1,2})$/.exec(yearText);
        if (yearMonth) {
            var month = parseInt(yearMonth[1], 10);
            if (month < 1 || month > 12) return false;
            return yearMonth[1].length === 2 || DATE_SEPARATOR.test(text.charAt(end));
        }

        if (!/^[12]\d{3}$/.test(yearText)) return false;
        var nextChar = findNonSpaceChar(text, end, 1);
        var prevChar = findNonSpaceChar(text, yearStart - 1, -1);
        return nextChar === "年" || DATE_SEPARATOR.test(nextChar) || DATE_SEPARATOR.test(prevChar);
    }

    /**
     * 郵便番号（〒 ddd-dddd）の3桁部分と4桁部分の開始位置を集める
     * @param {string} text - テキスト全体
     * @returns {object} 開始位置をキーにした true の表
     */
    function buildPostalCodeOffsets(text) {
        var postalOffsets = {};
        var postalPattern = /(〒\s*)?(\d{3})-(\d{4})/g;
        var halfWidthText = toHalfWidthDigitsHyphen(text);
        var match;
        while ((match = postalPattern.exec(halfWidthText)) !== null) {
            var firstPartOffset = match.index + (match[1] ? match[1].length : 0);
            postalOffsets[firstPartOffset] = true;      /* 3桁部分 / 3-digit part */
            postalOffsets[firstPartOffset + 4] = true;  /* ハイフンの後の4桁部分 / 4-digit part after the hyphen */
        }
        return postalOffsets;
    }

    /**
     * 郵便番号の一部か。位置の表を引き、なければ前後数文字に ddd-dddd があるかを見る
     * @param {object} postalOffsets - buildPostalCodeOffsets() の結果
     * @param {string} text - テキスト全体
     * @param {number} start - 数値の開始位置
     * @param {number} end - 数値の終了位置
     * @returns {boolean} 郵便番号の一部なら true
     */
    function isPostalCodePart(postalOffsets, text, start, end) {
        if (postalOffsets[start]) return true;
        var surroundingText = text.substring(start - 5, end + 6);
        return /(?:^|[^\d])(?:〒\s*)?\d{3}-\d{4}(?!\d)/.test(toHalfWidthDigitsHyphen(surroundingText));
    }

    /**
     * 電話番号の一部か。数字・区切り・括弧が続く範囲まで広げて照合する
     * @param {string} text - テキスト全体
     * @param {number} start - 数値の開始位置
     * @param {number} end - 数値の終了位置
     * @returns {boolean} 電話番号とみなせば true
     */
    function looksLikePhoneNumber(text, start, end) {
        var token = toHalfWidthForPhone(extractSurroundingToken(text, start, end, /[0-9０-９\-－ 　\(\)（）＋+]/));
        /* 「TEL 03-…」の前後の空白まで拾うと照合に失敗するので除く / Drop spaces picked up around the number, e.g. in "TEL 03-…" */
        token = token.replace(/^\s+|\s+$/g, "");
        /* 電話番号は 0・+・括弧で始まる。「12000 15000」のような数値の並びは電話番号にしない
           Phone numbers start with 0, + or a parenthesis; runs of plain numbers such as 12000 15000 are not phones */
        if (!/^[0+(]/.test(token)) return false;
        for (var i = 0; i < PHONE_PATTERNS.length; i++) {
            if (PHONE_PATTERNS[i].test(token)) return true;
        }
        return false;
    }

    /**
     * Luhn のチェックディジットが合うか
     * @param {string} digits - 数字だけの文字列
     * @returns {boolean} 合えば true
     */
    function passesLuhnCheck(digits) {
        var sum = 0;
        var doubleDigit = false;
        for (var i = digits.length - 1; i >= 0; i--) {
            var digitValue = digits.charCodeAt(i) - 48;
            if (doubleDigit) {
                digitValue *= 2;
                if (digitValue > 9) digitValue -= 9;
            }
            sum += digitValue;
            doubleDigit = !doubleDigit;
        }
        return sum % 10 === 0;
    }

    /**
     * クレジットカード番号の一部か。数字・ハイフン・スペース・括弧が続く範囲が13〜19桁で、Luhn チェックを通れば該当
     * @param {string} text - テキスト全体
     * @param {number} start - 数値の開始位置
     * @param {number} end - 数値の終了位置
     * @returns {boolean} カード番号とみなせば true
     */
    function looksLikeCreditCard(text, start, end) {
        var digits = toHalfWidthDigitsHyphen(extractSurroundingToken(text, start, end, /[0-9０-９\-－ \(\)]/)).replace(/\D+/g, "");
        return digits.length >= 13 && digits.length <= 19 && passesLuhnCheck(digits);
    }

    /**
     * コロンかハイフンで区切った MAC アドレスの一部か
     * @param {string} text - テキスト全体
     * @param {number} start - 数値の開始位置
     * @param {number} end - 数値の終了位置
     * @returns {boolean} MAC アドレスとみなせば true
     */
    function looksLikeMacAddress(text, start, end) {
        var token = toHalfWidthForMac(extractSurroundingToken(text, start, end, /[0-9０-９A-Fa-fＡ-Ｆａ-ｆ:\-：－]/));
        return /^(?:[0-9A-Fa-f]{2}:){5}[0-9A-Fa-f]{2}$/.test(token) || /^(?:[0-9A-Fa-f]{2}-){5}[0-9A-Fa-f]{2}$/.test(token);
    }

    /**
     * 除外ルールに当たらないか
     * @param {string} token - 元のテキストにあるままの数値
     * @param {number} start - 数値の開始位置
     * @param {string} text - テキスト全体
     * @param {object} exclusionOptions - 除外ルールのキーと有効／無効
     * @param {object|null} postalOffsets - 郵便番号の位置の表。郵便番号を除外しないときは null
     * @returns {boolean} カンマを付けてよければ true
     */
    function isEligibleNumber(token, start, text, exclusionOptions, postalOffsets) {
        var end = start + token.length;
        var value = toHalfWidthNumber(removeCommas(token));
        /* 0で始まる数値と、英字・# に続く数値は番号やコードなので、設定に関係なく除外する
           Numbers starting with 0 or following letters or # are codes, excluded regardless of settings */
        if (hasLeadingZero(value)) return false;
        if (isPartOfCode(text, start)) return false;
        if (postalOffsets && isPostalCodePart(postalOffsets, text, start, end)) return false;
        if (exclusionOptions.slashAdjacent && isNextToSlash(text, start, end)) return false;
        if (exclusionOptions.years && looksLikeYear(value, text, start, end)) return false;
        if (exclusionOptions.phoneNumbers && looksLikePhoneNumber(text, start, end)) return false;
        if (exclusionOptions.creditCards && looksLikeCreditCard(text, start, end)) return false;
        if (exclusionOptions.macAddresses && looksLikeMacAddress(text, start, end)) return false;
        return true;
    }

    // =========================================
    // テキストの収集 / Collecting text frames
    // =========================================

    /**
     * アイテムからテキストフレームを集める。グループ（クリップグループを含む）は中へたどる
     * @param {PageItem} pageItem - 対象のアイテム
     * @param {TextFrame[]} textFrames - 集めたテキストフレームを入れる配列
     * @returns {void}
     */
    function collectTextFramesFromItem(pageItem, textFrames) {
        if (!pageItem) return;
        if (pageItem.typename === "TextFrame") {
            textFrames.push(pageItem);
            return;
        }
        if (pageItem.typename !== "GroupItem") return;

        /* 直下のテキストを集め、入れ子のグループへ再帰する / Take direct text frames, then recurse into nested groups */
        for (var i = 0; i < pageItem.textFrames.length; i++) {
            textFrames.push(pageItem.textFrames[i]);
        }
        for (var j = 0; j < pageItem.groupItems.length; j++) {
            collectTextFramesFromItem(pageItem.groupItems[j], textFrames);
        }
    }

    /**
     * 対象範囲のテキストフレームを集める
     * @param {Document} doc - 対象のドキュメント
     * @param {string} targetScope - "selection" または "document"
     * @returns {TextFrame[]} テキストフレームの配列
     */
    function collectTargetTextFrames(doc, targetScope) {
        var textFrames = [];
        if (targetScope === "document") {
            /* pageItems はグループ内も並べるので、たどると二重になる。textFrames なら1回ずつ
               pageItems also lists grouped items, so walking it doubles them; textFrames lists each once */
            var documentTextFrames = doc.textFrames;
            for (var i = 0; i < documentTextFrames.length; i++) {
                textFrames.push(documentTextFrames[i]);
            }
            return textFrames;
        }
        var selectedItems = doc.selection;
        for (var j = 0; j < selectedItems.length; j++) {
            collectTextFramesFromItem(selectedItems[j], textFrames);
        }
        return textFrames;
    }

    /**
     * テキストフレームの番号を、上から下（同じ高さは左から右）の順に並べる
     * @param {TextFrame[]} textFrames - テキストフレームの配列
     * @returns {number[]} 並べ替えた配列の番号
     */
    function sortFrameIndexesByPosition(textFrames) {
        var positions = [];
        var frameIndexes = [];
        for (var i = 0; i < textFrames.length; i++) {
            var frameBounds = textFrames[i].geometricBounds;
            positions.push({ left: frameBounds[0], top: frameBounds[1] });
            frameIndexes.push(i);
        }
        frameIndexes.sort(function (indexA, indexB) {
            var topA = positions[indexA].top;
            var topB = positions[indexB].top;
            var topOrder = RANK_FROM_TOP ? (topB - topA) : (topA - topB);
            if (topOrder !== 0) return topOrder;
            return positions[indexA].left - positions[indexB].left;
        });
        return frameIndexes;
    }

    // =========================================
    // 候補の検出 / Finding candidates
    // =========================================

    /* 数値のトークン。カンマ入り（位置の誤りも含む）を優先し、なければカンマなしの数字の並び
       Number tokens: comma-grouped ones first (misplaced commas included), otherwise plain digit runs */
    var NUMBER_TOKEN_PATTERN = /[+\-＋－]?(?:[0-9０-９]{1,3}(?:[，,][0-9０-９]+)+|[0-9０-９]+)(?:[\.．][0-9０-９]+)?/g;

    /**
     * 1つのテキストから、カンマを付ける（直す）数値を集める
     * @param {string} text - テキストフレームの内容
     * @param {object} exclusionOptions - 除外ルールのキーと有効／無効
     * @returns {object[]} 候補の配列 { token, offset, newText }
     */
    function collectFrameCandidates(text, exclusionOptions) {
        var postalOffsets = exclusionOptions.postalCodes ? buildPostalCodeOffsets(text) : null;
        var candidates = [];
        var match;
        NUMBER_TOKEN_PATTERN.lastIndex = 0;
        while ((match = NUMBER_TOKEN_PATTERN.exec(text)) !== null) {
            var token = match[0];
            if (!needsCommaFix(token)) continue;
            if (!isEligibleNumber(token, match.index, text, exclusionOptions, postalOffsets)) continue;
            candidates.push({ token: token, offset: match.index, newText: formatNumberWithCommas(token) });
        }
        return candidates;
    }

    /**
     * プレビューに並べる行を、テキストの位置順に作る
     * @param {TextFrame[]} textFrames - テキストフレームの配列
     * @param {object} exclusionOptions - 除外ルールのキーと有効／無効
     * @returns {object[]} 行の配列 { token, offset, newText, frameIndex, frameNumber }
     */
    function buildPreviewEntries(textFrames, exclusionOptions) {
        var previewEntries = [];
        var orderedIndexes = sortFrameIndexesByPosition(textFrames);
        for (var rank = 0; rank < orderedIndexes.length; rank++) {
            var frameIndex = orderedIndexes[rank];
            var candidates = collectFrameCandidates(textFrames[frameIndex].contents, exclusionOptions);
            for (var i = 0; i < candidates.length; i++) {
                candidates[i].frameIndex = frameIndex;
                candidates[i].frameNumber = rank + 1;
                previewEntries.push(candidates[i]);
            }
        }
        return previewEntries;
    }

    /**
     * 行をテキストフレームごとにまとめる
     * @param {object[]} entries - プレビューの行
     * @returns {object} フレーム番号をキーにした行の配列
     */
    function groupEntriesByFrame(entries) {
        var entriesByFrame = {};
        for (var i = 0; i < entries.length; i++) {
            var frameIndex = entries[i].frameIndex;
            if (!entriesByFrame[frameIndex]) entriesByFrame[frameIndex] = [];
            entriesByFrame[frameIndex].push(entries[i]);
        }
        return entriesByFrame;
    }

    // =========================================
    // テキストの書き換え / Rewriting text
    // =========================================

    /**
     * 控えた位置に、同じ数値がまだあるか（先頭と末尾の文字で確かめる）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} offset - 数値の開始位置
     * @param {string} token - 数値
     * @returns {boolean} 一致すれば true
     */
    function isTokenAt(textFrame, offset, token) {
        var lastCharIndex = offset + token.length - 1;
        if (lastCharIndex >= textFrame.characters.length) return false;
        return textFrame.characters[offset].contents === token.charAt(0) &&
            textFrame.characters[lastCharIndex].contents === token.charAt(token.length - 1);
    }

    /**
     * 数値の文字を、カンマの出し入れと記号の半角／全角の置き換えだけで書き換える。
     * 数字の文字には触れないので、その書式は残る。右から処理して位置をずらさない
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} offset - 数値の開始位置
     * @param {string} oldText - 今の数値
     * @param {string} newText - カンマを付けた数値
     * @returns {void}
     */
    function rewriteNumberCharacters(textFrame, offset, oldText, newText) {
        var oldIndex = oldText.length - 1;
        var newIndex = newText.length - 1;
        while (oldIndex >= 0) {
            var oldChar = oldText.charAt(oldIndex);
            var newChar = newText.charAt(newIndex);
            if (oldChar === newChar) {
                oldIndex--;
                newIndex--;
            } else if (isComma(newChar) && !isComma(oldChar)) {
                /* 数字の右にカンマを足す（カンマは数字の書式を引き継ぐ）/ Add a comma after the digit; it inherits the digit's formatting */
                textFrame.characters[offset + oldIndex].contents = oldChar + newChar;
                newIndex--;
            } else if (isComma(oldChar) && !isComma(newChar)) {
                /* 位置の誤ったカンマを消す / Remove a misplaced comma */
                textFrame.characters[offset + oldIndex].remove();
                oldIndex--;
            } else {
                /* 符号・カンマ・小数点の半角／全角を数字にそろえる / Match the width of a sign, comma or point to the digits */
                textFrame.characters[offset + oldIndex].contents = newChar;
                oldIndex--;
                newIndex--;
            }
        }
    }

    /**
     * テキストフレーム内の選ばれた数値にカンマを付ける
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {object[]} frameEntries - このフレームの選ばれた行
     * @returns {void}
     */
    function applyEntriesToFrame(textFrame, frameEntries) {
        /* 右の数値から書き換え、左の数値の位置をずらさない / Rewrite from the rightmost number so offsets on the left stay valid */
        frameEntries.sort(function (entryA, entryB) {
            return entryB.offset - entryA.offset;
        });
        for (var i = 0; i < frameEntries.length; i++) {
            var entry = frameEntries[i];
            if (!isTokenAt(textFrame, entry.offset, entry.token)) continue;
            rewriteNumberCharacters(textFrame, entry.offset, entry.token, entry.newText);
        }
    }

    // =========================================
    // ダイアログ / Dialogs
    // =========================================

    /**
     * 不透明度と表示位置をそろえたダイアログを作る
     * @param {string} title - ダイアログのタイトル
     * @returns {Window} ダイアログ
     */
    function createDialog(title) {
        var dialogWindow = new Window("dialog", title);
        dialogWindow.opacity = DIALOG_OPACITY;
        /* 既定の位置からずらして表示する / Shift the dialog from its default position */
        dialogWindow.onShow = function () {
            dialogWindow.location = [dialogWindow.location[0] + DIALOG_OFFSET_X, dialogWindow.location[1] + DIALOG_OFFSET_Y];
        };
        return dialogWindow;
    }

    /**
     * 縦並び・左揃えのパネルを追加する
     * @param {Window} parent - 追加先
     * @param {object} labelSet - パネル名のラベル定義
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, labelSet) {
        var newPanel = parent.add("panel", undefined, getLabel(labelSet));
        newPanel.orientation = "column";
        newPanel.alignChildren = "left";
        newPanel.margins = PANEL_MARGINS;
        return newPanel;
    }

    /**
     * 中央揃えのボタン行を追加する
     * @param {Window} dialogWindow - 追加先のダイアログ
     * @returns {Group} ボタン行
     */
    function addButtonRow(dialogWindow) {
        var btnRowGroup = dialogWindow.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "center";
        return btnRowGroup;
    }

    /**
     * ［キャンセル］［OK］ボタンを追加する
     * @param {Group} btnRowGroup - 追加先のボタン行
     * @returns {void}
     */
    function addCancelOkButtons(btnRowGroup) {
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
    }

    /**
     * 対象範囲と除外ルールを選ぶダイアログを出す
     * @returns {object|null} { targetScope, exclusionOptions }。キャンセルなら null
     */
    function showScopeDialog() {
        var scopeDialog = createDialog(getLabel(LABELS.dialog.scopeTitle) + " " + SCRIPT_VERSION);
        scopeDialog.alignChildren = "fill";

        var scopePanel = addPanel(scopeDialog, LABELS.panel.scope);
        var selectionRadio = scopePanel.add("radiobutton", undefined, getLabel(LABELS.radio.selection));
        scopePanel.add("radiobutton", undefined, getLabel(LABELS.radio.wholeDocument));
        selectionRadio.value = true;

        var exclusionPanel = addPanel(scopeDialog, LABELS.panel.exclusion);
        var exclusionCheckboxes = [];
        for (var i = 0; i < EXCLUSION_RULES.length; i++) {
            var ruleKey = EXCLUSION_RULES[i].key;
            var exclusionCheckbox = exclusionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox[ruleKey]));
            exclusionCheckbox.value = EXCLUSION_RULES[i].defaultValue;
            exclusionCheckbox.helpTip = getLabel(LABELS.tooltip[ruleKey]);
            exclusionCheckboxes.push(exclusionCheckbox);
        }

        addCancelOkButtons(addButtonRow(scopeDialog));

        if (scopeDialog.show() !== 1) return null;

        var exclusionOptions = {};
        for (var j = 0; j < EXCLUSION_RULES.length; j++) {
            exclusionOptions[EXCLUSION_RULES[j].key] = exclusionCheckboxes[j].value;
        }
        return {
            targetScope: selectionRadio.value ? "selection" : "document",
            exclusionOptions: exclusionOptions
        };
    }

    /**
     * 行数からプレビューリストの高さを決める
     * @param {number} rowCount - 行数
     * @returns {number} リストの高さ
     */
    function calcPreviewListHeight(rowCount) {
        var visibleRows = Math.max(PREVIEW_MIN_ROWS, Math.min(rowCount, PREVIEW_MAX_ROWS));
        return Math.min(PREVIEW_HEADER_HEIGHT + PREVIEW_ROW_HEIGHT * visibleRows + PREVIEW_LIST_PADDING, PREVIEW_LIST_MAX_HEIGHT);
    }

    /**
     * 行に ✓ が付いているか
     * @param {ListItem} listRow - プレビューリストの行
     * @returns {boolean} ✓ が付いていれば true
     */
    function isRowChecked(listRow) {
        return listRow.text === CHECK_MARK;
    }

    /**
     * すべての行に ✓ を付ける
     * @param {ListBox} previewList - プレビューリスト
     * @returns {void}
     */
    function checkAllRows(previewList) {
        for (var i = 0; i < previewList.items.length; i++) {
            previewList.items[i].text = CHECK_MARK;
        }
    }

    /**
     * カンマを付ける数値を選ぶプレビューダイアログを出す
     * @param {object[]} previewEntries - buildPreviewEntries() の結果
     * @returns {object[]|null} ✓ の付いた行。キャンセルなら null
     */
    function showPreviewDialog(previewEntries) {
        var previewDialog = createDialog(getLabel(LABELS.dialog.previewTitle));
        previewDialog.orientation = "column";
        previewDialog.alignChildren = "fill";

        var instructionText = previewDialog.add("statictext", undefined, getLabel(LABELS.message.previewInstruction));
        instructionText.alignment = "fill";

        var previewList = previewDialog.add("listbox", undefined, [], {
            multiselect: false,
            numberOfColumns: 3,
            showHeaders: true,
            columnTitles: [getLabel(LABELS.listColumn.select), getLabel(LABELS.listColumn.frameNumber), getLabel(LABELS.listColumn.value)]
        });
        previewList.preferredSize = [PREVIEW_LIST_WIDTH, calcPreviewListHeight(previewEntries.length)];
        previewList.columnWidths = PREVIEW_COLUMN_WIDTHS;
        previewList.helpTip = getLabel(LABELS.tooltip.previewList);
        for (var i = 0; i < previewEntries.length; i++) {
            var listRow = previewList.add("item", CHECK_MARK);
            listRow.subItems[0].text = "#" + previewEntries[i].frameNumber;
            listRow.subItems[1].text = previewEntries[i].token;
            listRow.helpTip = previewEntries[i].token;
        }
        /* クリックした行の ✓ だけを付け外しし、選択は解除して同じ行を続けて押せるようにする
           Toggle only the clicked row, then clear the selection so the same row can be clicked again */
        previewList.onChange = function () {
            var clickedRow = previewList.selection;
            if (!clickedRow) return;
            clickedRow.text = isRowChecked(clickedRow) ? "" : CHECK_MARK;
            previewList.selection = null;
        };

        var btnRowGroup = addButtonRow(previewDialog);
        var btnSelectAll = btnRowGroup.add("button", undefined, getLabel(LABELS.button.selectAll));
        btnSelectAll.helpTip = getLabel(LABELS.tooltip.selectAll);
        btnSelectAll.onClick = function () {
            checkAllRows(previewList);
        };
        addCancelOkButtons(btnRowGroup);

        if (previewDialog.show() !== 1) return null;

        var chosenEntries = [];
        for (var j = 0; j < previewList.items.length; j++) {
            if (isRowChecked(previewList.items[j])) chosenEntries.push(previewEntries[j]);
        }
        return chosenEntries;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで対象と除外ルールを選び、プレビューで確認してからカンマを付ける
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var dialogResult = showScopeDialog();
        if (!dialogResult) return;

        var textFrames = collectTargetTextFrames(app.activeDocument, dialogResult.targetScope);
        if (textFrames.length === 0) {
            alert(getLabel(LABELS.alert.noTextFrames));
            return;
        }

        var previewEntries = buildPreviewEntries(textFrames, dialogResult.exclusionOptions);
        if (previewEntries.length === 0) {
            alert(getLabel(LABELS.alert.noCandidates));
            return;
        }

        var chosenEntries = showPreviewDialog(previewEntries);
        if (!chosenEntries) return;

        var entriesByFrame = groupEntriesByFrame(chosenEntries);
        for (var i = 0; i < textFrames.length; i++) {
            if (!entriesByFrame[i]) continue;
            /* ロックされたテキストなどは書き換えで例外になるので飛ばす / Skip text that throws on editing, such as locked text */
            try {
                applyEntriesToFrame(textFrames[i], entriesByFrame[i]);
            } catch (e) {}
        }
    }

    main();

})();
