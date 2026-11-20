#target illustrator
#targetengine "FormatNumberWithCommasEngine"
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.FormatNumberWithCommas
Name=数字に桁区切りのカンマを付ける
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.12
Release-Date=2025-08-12
Description-BEGIN
テキスト内の数値に桁区切りのカンマを付け、位置の誤ったカンマを直します。
西暦・郵便番号・電話番号などは除外し、書き換える数値はプレビューで選べます。

Adds thousands separators to numbers in text and fixes misplaced commas.
Years, postal codes, phone numbers and the like are skipped, and you pick the numbers to change in a preview.
Description-END
SCRIPTMETA-END
*/

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
var SCRIPT_VERSION  = "v1.0.12";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-12";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-20";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "数字に桁区切りのカンマを付ける"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FormatNumberWithCommas.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FormatNumberWithCommas.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n21f07978f177"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/FormatNumberWithCommas.jsx"; /* 最新版のソース / latest source */

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
    var PREVIEW_LIST_WIDTH = 260;               /* プレビューリストの幅 / preview list width */
    var PREVIEW_COLUMN_WIDTHS = [40, 70, 150];  /* 列幅 [選択, #, 数値] / column widths */
    var PREVIEW_HEADER_HEIGHT = 24;             /* 見出し行の高さ / header row height */
    var PREVIEW_ROW_HEIGHT = 20;                /* 1行の高さ / row height */
    var PREVIEW_LIST_PADDING = 40;              /* リストの上下の余白 / list padding */
    var PREVIEW_LIST_MAX_HEIGHT = 520;          /* リストの最大の高さ / max list height */
    var PREVIEW_MIN_ROWS = 4;                   /* 最低限見せる行数 / minimum visible rows */
    var PREVIEW_MAX_ROWS = 16;                  /* スクロールせずに見せる行数 / rows shown before scrolling */
    var CHECK_MARK = "✓";                       /* カンマを付ける行の印 / mark on rows that get commas */

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

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

    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)

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

    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds

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
        /* グループ（クリップグループを含む）は中へたどる。文字の選択は対象外
           Walk into groups (clip groups included); text selections are not used */
        return collectSelectionTextFrames(doc.selection, { textRangeToFrame: false, unique: false });
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

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

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
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // =========================================
    // ダイアログ / Dialogs
    // =========================================

    /**
     * ラジオやチェックボックスを縦に並べるパネルを追加する
     * @param {Window} parent - 追加先
     * @param {object} labelSet - パネル名のラベル定義
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, labelSet) {
        var newPanel = parent.add("panel", undefined, getLabel(labelSet));
        setupPanel(newPanel, 6);
        return newPanel;
    }

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
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

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // UI の明暗（再利用パーツ） / UI theme (reusable)

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
       9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
       / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
         nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
    var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
    var VERSION_FOOTER_GLYPHS = {
        "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
        "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
        "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
        "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
        "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
        "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
        "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
        "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
        "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
        "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
        "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
        "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
        "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
        "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
        "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
        "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
        "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
        "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
        "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
        "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
        "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
        "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
        "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
        "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
        "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
        "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
        "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
        "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
        "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
        "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
        "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
        "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
        "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
        "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
        "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
        "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
        "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
        "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
        "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
        ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
        "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
        "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
        "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
        " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
    };

    /**
     * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPT_VERSION）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

        /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
           左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
           / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
             or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
        var versionText = String(SCRIPT_VERSION);
        var leftCandidates = [versionText];
        if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
        var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
        versionCanvas.alignment = ["fill", "center"];
        bindVersionInfoClick(versionCanvas); /* クリックで「このスクリプトについて」を開く / click opens About This Script */
        var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
        brandCanvas.alignment = ["right", "center"];
        brandCanvas.maximumSize = brandCanvas.preferredSize;
        targetWindow.versionFooter = footerGroup;
        return versionCanvas;
    }

    /**
     * 文字列のすべての文字に字形があるかを返す
     * @param {string} dotText - 文字列
     * @returns {boolean} すべてあれば true
     */
    function hasVersionFooterGlyphs(dotText) {
        for (var i = 0; i < dotText.length; i++) {
            if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
        }
        return dotText.length > 0;
    }

    /**
     * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
     * @param {Group} parentGroup - 足す先
     * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
     * @returns {Group} 描くグループ
     */
    function addVersionFooterCanvas(parentGroup, dotTexts) {
        var dotCanvas = parentGroup.add("group");
        var shortestText = dotTexts[dotTexts.length - 1];
        dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
        dotCanvas.helpTip = dotTexts[0];
        dotCanvas.onDraw = function () {
            var availableWidth = this.size ? this.size.width : 0;
            var drawText = shortestText;
            for (var i = 0; i < dotTexts.length; i++) {
                if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                    drawText = dotTexts[i];
                    break;
                }
            }
            drawVersionFooterDots(this.graphics, drawText);
        };
        return dotCanvas;
    }

    /**
     * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - 文字列
     * @returns {number} 幅（ドット）
     */
    function measureVersionFooterDots(versionText) {
        var dotCount = 0;
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
            dotCount += glyphRows[0].length;
        }
        return dotCount;
    }

    /**
     * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - 文字列
     * @returns {void}
     */
    function drawVersionFooterDots(canvasGraphics, versionText) {
        var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
        var dotSize = VERSION_FOOTER_PIXEL_SIZE;
        var cursorDots = 0;
        canvasGraphics.newPath();
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
            for (var row = 0; row < glyphRows.length; row++) {
                var rowText = glyphRows[row];
                var runStart = -1;
                for (var col = 0; col <= rowText.length; col++) {
                    var isDot = col < rowText.length && rowText.charAt(col) === "#";
                    if (isDot && runStart < 0) runStart = col;
                    if (!isDot && runStart >= 0) {
                        canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                        runStart = -1;
                    }
                }
            }
            cursorDots += glyphRows[0].length;
        }
        canvasGraphics.fillPath(footerBrush);
    }

    // -----------------------------------------
    // このスクリプトについて / About this script
    // -----------------------------------------
    /* 左下のスクリプト名・バージョンをクリックすると開く。基本情報ブロックの SCRIPT_* を読み、
       この部品だけで動くよう、文言（VERSION_INFO_LABELS）とレイアウトの値も部品の中に持つ
       Opened by clicking the name and version at the bottom left. Reads the SCRIPT_* basic info,
       and keeps its own labels and layout values so the part works on its own */

    var VERSION_INFO_LANG = (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en"; /* 文言の言語 / label language */
    var VERSION_INFO_IS_INDESIGN = (String(app.name).indexOf("InDesign") >= 0);         /* InDesign で動いているか / running in InDesign */
    /* 実行中のスクリプトのファイル。ボタンを押すころには $.fileName が別のものを指すことがあるので、読み込んだ時点で控える
       The running script file, captured at load time; $.fileName may point elsewhere by the time the button is clicked */
    var VERSION_INFO_SCRIPT_FILE = new File($.fileName);
    var VERSION_INFO_REVEAL_APP_PATH = "/Applications/RevealInFinder.app";   /* Finder で選択表示する補助アプリ / helper app that reveals a file in the Finder */
    var VERSION_INFO_REVEAL_PATH_FILE = "/tmp/illustrator_reveal_path.txt";  /* 補助アプリが読むパス（アプリの旧名のまま。変えると読めない）/ path the helper reads (keeps the app's old name) */

    var VERSION_INFO_WINDOW_MARGINS = 16;               /* ウィンドウの余白 / window margin */
    var VERSION_INFO_WINDOW_SPACING = 12;               /* ウィンドウ内の間隔 / window spacing */
    var VERSION_INFO_PANEL_MARGINS = [16, 20, 16, 12];  /* パネルの余白 [左,上,右,下] / panel margins */
    var VERSION_INFO_ROW_SPACING = 6;                   /* 項目名と値の間隔 / gap between a label and its value */
    var VERSION_INFO_BUTTON_SPACING = 10;               /* ボタンどうしの間隔 / gap between buttons */
    var VERSION_INFO_BUTTON_ROW_MARGINS = [0, 5, 0, 14]; /* ボタン行の余白（Illustrator 標準のダイアログに合わせる）/ button row margins */
    var VERSION_INFO_COPY_BUTTON_SIZE = 22;             /* コピーボタンの一辺 / copy button size */
    /* コピーボタンの配色。∧∨（ステップボタン）と同じく、背景に重ねる半透明 / Copy button colors, translucent overlays like the stepper */
    var VERSION_INFO_COPY_FILL_COLOR    = isDarkUI() ? [0, 0, 0, 0.10] : [1, 1, 1, 0.50]; /* 地 / background */
    var VERSION_INFO_COPY_FRAME_COLOR   = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.15]; /* 枠線 / frame */
    var VERSION_INFO_COPY_PRESSED_COLOR = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 押下中 / pressed */
    var VERSION_INFO_COPY_ICON_COLOR    = isDarkUI() ? [1, 1, 1, 0.85] : [0, 0, 0, 0.70]; /* アイコン / icon */

    var VERSION_INFO_LABELS = {
        dialog: {
            title: { ja: "このスクリプトについて", en: "About This Script" }
        },
        panel: {
            generalInfo: { ja: "基本情報", en: "General" },
            links: { ja: "リンク", en: "Links" }
        },
        fieldLabel: {
            scriptName: { ja: "スクリプト名", en: "Script Name" },
            scriptFeature: { ja: "機能", en: "Feature" },
            scriptVersion: { ja: "バージョン", en: "Version" }
        },
        button: {
            openReadme: { ja: "README", en: "README" },
            openSourceCode: { ja: "ソースコード", en: "Source Code" },
            openArticle: { ja: "記事（note）", en: "Article (note)" },
            revealInFinder: { ja: "Finder に表示", en: "Show in Finder" },
            revealInExplorer: { ja: "エクスプローラーに表示", en: "Show in Explorer" },
            close: { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            versionInfo: { ja: "クリックで、このスクリプトについての情報を表示", en: "Click for info about this script" },
            copyScriptName: { ja: "スクリプト名をコピー", en: "Copy Script Name" }
        },
        alert: {
            copyFailed: { ja: "スクリプト名をコピーできませんでした。", en: "Could not copy the script name." }
        }
    };

    /**
     * VERSION_INFO_LABELS の文言を今の言語で返す
     * @param {Object} labelEntry - { ja, en }
     * @returns {string} 文言
     */
    function getVersionInfoLabel(labelEntry) {
        return labelEntry[VERSION_INFO_LANG] || labelEntry.en;
    }

    /**
     * 基本情報ブロックの値を返す（変数が無い・文字列でなければ空文字）。typeof で読むので、未定義でも例外にならない
     * @param {string} variableKey - "name" / "version" / "feature" / "readme" / "source" / "article"
     * @returns {string} 値
     */
    function getVersionInfoValue(variableKey) {
        var infoValue;
        if (variableKey === "name") infoValue = (typeof SCRIPT_NAME !== "undefined") ? SCRIPT_NAME : "";
        else if (variableKey === "version") infoValue = (typeof SCRIPT_VERSION !== "undefined") ? SCRIPT_VERSION : "";
        else if (variableKey === "feature") infoValue = (typeof SCRIPT_DESCRIPTION_JA !== "undefined") ? SCRIPT_DESCRIPTION_JA : "";
        else if (variableKey === "readme") {
            if (VERSION_INFO_LANG === "ja") infoValue = (typeof SCRIPT_README_JA !== "undefined") ? SCRIPT_README_JA : "";
            else infoValue = (typeof SCRIPT_README_EN !== "undefined") ? SCRIPT_README_EN : "";
        }
        else if (variableKey === "source") infoValue = (typeof SCRIPT_SOURCE_URL !== "undefined") ? SCRIPT_SOURCE_URL : "";
        else if (variableKey === "article") infoValue = (typeof SCRIPT_ARTICLE_URL !== "undefined") ? SCRIPT_ARTICLE_URL : "";
        return (infoValue === undefined || infoValue === null) ? "" : String(infoValue);
    }

    /**
     * 左下のスクリプト名・バージョンをクリックしたら、「このスクリプトについて」のダイアログを開くようにする
     * @param {Group} versionCanvas - addVersionFooterCanvas() で作った左の欄
     * @returns {void}
     */
    function bindVersionInfoClick(versionCanvas) {
        versionCanvas.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.versionInfo);
        /* mousedown で開くと、離したときのイベントが閉じたあとの欄に届くので mouseup で開く
           Open on mouseup so the release does not land on the footer after the dialog closes */
        versionCanvas.addEventListener("mouseup", function () {
            showVersionInfoDialog(versionCanvas.window.type === "palette");
        });
    }

    /**
     * 「このスクリプトについて」のダイアログを開く。スクリプト名（コピーボタン付き）・機能・バージョンの［基本情報］パネル、
     * README・ソースコード・記事を開く［リンク］パネル、［Finder に表示］［閉じる］を並べる
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true（Illustrator のコピーを BridgeTalk 経由にする）
     * @returns {void}
     */
    function showVersionInfoDialog(isFromPalette) {
        var infoDialog = new Window("dialog", getVersionInfoLabel(VERSION_INFO_LABELS.dialog.title));
        infoDialog.orientation = "column";
        infoDialog.alignChildren = "fill";
        infoDialog.margins = VERSION_INFO_WINDOW_MARGINS;
        infoDialog.spacing = VERSION_INFO_WINDOW_SPACING;

        var infoPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.generalInfo);
        /* コピーボタンはスクリプト名だけ / Only the script name gets a copy button */
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptName, getVersionInfoValue("name"), isFromPalette);
        /* 機能は README の見出しで日本語だけなので、日本語の画面だけに出す / The feature is the Japanese README title, shown only in the Japanese UI */
        if (VERSION_INFO_LANG === "ja" && getVersionInfoValue("feature") !== "") {
            addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptFeature, getVersionInfoValue("feature"), null);
        }
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptVersion, getVersionInfoValue("version"), null);
        alignVersionInfoLabels(infoPanel);

        /* URL が無いボタンはディム表示にして、並びはスクリプトごとに変えない / Dim buttons without a URL; keep the same buttons everywhere */
        var linksPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.links);
        var linkButtonRow = linksPanel.add("group");
        linkButtonRow.orientation = "row";
        linkButtonRow.alignment = ["left", "center"];
        linkButtonRow.alignChildren = ["left", "center"];
        linkButtonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openReadme, getVersionInfoValue("readme"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openSourceCode, getVersionInfoValue("source"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openArticle, getVersionInfoValue("article"));

        /* ボタン行：左に［Finder に表示］、右に［閉じる］ / Button row: Show in Finder at the left, Close at the right */
        var buttonRow = infoDialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = ["fill", "bottom"];
        buttonRow.alignChildren = ["left", "center"];
        buttonRow.margins = VERSION_INFO_BUTTON_ROW_MARGINS;
        buttonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        var revealLabel = (File.fs === "Windows") ? VERSION_INFO_LABELS.button.revealInExplorer : VERSION_INFO_LABELS.button.revealInFinder;
        var revealButton = buttonRow.add("button", undefined, getVersionInfoLabel(revealLabel));
        if (VERSION_INFO_SCRIPT_FILE.exists) {
            revealButton.helpTip = VERSION_INFO_SCRIPT_FILE.fsName;
            revealButton.onClick = function () {
                revealVersionInfoScriptFile(VERSION_INFO_SCRIPT_FILE);
                infoDialog.close();
            };
        } else {
            revealButton.enabled = false;
        }
        /* 余りの幅を吸って、［閉じる］を右端に寄せる / Absorbs the extra width so Close sits at the right edge */
        var buttonSpacer = buttonRow.add("group");
        buttonSpacer.alignment = ["fill", "fill"];
        buttonSpacer.minimumSize.width = 0;
        var closeButton = buttonRow.add("button", undefined, getVersionInfoLabel(VERSION_INFO_LABELS.button.close), { name: "ok" });
        closeButton.alignment = ["right", "center"];
        /* return でも esc でも閉じる / Close with both Return and Esc */
        infoDialog.defaultElement = closeButton;
        infoDialog.cancelElement = closeButton;
        infoDialog.onShow = function () {
            pinVersionInfoCopyButton(infoPanel);
        };
        infoDialog.show();
    }

    /**
     * 「このスクリプトについて」のダイアログにパネルを足す
     * @param {Window} infoDialog - 足す先のダイアログ
     * @param {Object} titleEntry - パネル名の文言 { ja, en }
     * @returns {Panel} パネル
     */
    function addVersionInfoPanel(infoDialog, titleEntry) {
        var infoPanel = infoDialog.add("panel", undefined, getVersionInfoLabel(titleEntry));
        infoPanel.orientation = "column";
        infoPanel.alignChildren = ["fill", "top"];
        infoPanel.alignment = "fill";
        infoPanel.margins = VERSION_INFO_PANEL_MARGINS;
        infoPanel.spacing = VERSION_INFO_WINDOW_SPACING;
        return infoPanel;
    }

    /**
     * ［基本情報］パネルに「項目名 : 値」の1行を足す（値は見せるだけなので statictext）
     * @param {Panel} infoPanel - 足す先のパネル
     * @param {Object} labelEntry - 項目名の文言 { ja, en }
     * @param {string} valueText - 表示する値
     * @param {boolean|null} copyFromPalette - 値の右にコピーボタンを置くなら true / false（常駐パレットから開いたか）、置かないなら null
     * @returns {StaticText} 値の statictext
     */
    function addVersionInfoRow(infoPanel, labelEntry, valueText, copyFromPalette) {
        var infoRow = infoPanel.add("group");
        infoRow.orientation = "row";
        infoRow.alignment = ["fill", "center"];
        infoRow.alignChildren = ["left", "center"];
        infoRow.spacing = VERSION_INFO_ROW_SPACING;
        /* 項目名のコロンは日本語が「 :」、英語が「:」 / Colon: " :" in Japanese, ":" in English */
        var rowLabel = infoRow.add("statictext", undefined, getVersionInfoLabel(labelEntry) + (VERSION_INFO_LANG === "ja" ? " :" : ":"));
        rowLabel.justify = "right";
        var valueLabel = infoRow.add("statictext", undefined, valueText);
        if (copyFromPalette === null) return valueLabel;
        /* statictext の fill では伸びないことがあるので、伸びるスペーサーでコピーボタンをパネルの右端に寄せる
           A stretching spacer (statictext fill is unreliable) pushes the copy button to the panel's right edge */
        var copySpacer = infoRow.add("group");
        copySpacer.alignment = ["fill", "fill"];
        copySpacer.minimumSize.width = 0;
        infoRow.copyButton = addVersionInfoCopyButton(infoRow, valueLabel, copyFromPalette);
        return valueLabel;
    }

    /**
     * コピーボタンを［基本情報］パネルの右端に置き直す（表示時、レイアウトが決まったあとに呼ぶ）。
     * 行の fill・スペーサーだけでは、項目の値が短い行がパネルの幅まで伸びないことがあるので、行の幅とボタンの位置を直接入れる
     * @param {Panel} infoPanel - ［基本情報］パネル
     * @returns {void}
     */
    function pinVersionInfoCopyButton(infoPanel) {
        var innerWidth = infoPanel.size.width - VERSION_INFO_PANEL_MARGINS[0] - VERSION_INFO_PANEL_MARGINS[2];
        for (var i = 0; i < infoPanel.children.length; i++) {
            var infoRow = infoPanel.children[i];
            if (!infoRow.copyButton || !infoRow.size || infoRow.size.width >= innerWidth) continue;
            infoRow.size = [innerWidth, infoRow.size.height];
            infoRow.copyButton.location = [innerWidth - infoRow.copyButton.size.width, infoRow.copyButton.location[1]];
        }
    }

    /**
     * ［基本情報］パネルの項目名を、いちばん長いものの幅にそろえる
     * @param {Panel} infoPanel - 対象のパネル
     * @returns {void}
     */
    function alignVersionInfoLabels(infoPanel) {
        var rowLabels = [];
        var labelWidth = 0;
        for (var i = 0; i < infoPanel.children.length; i++) {
            var rowLabel = infoPanel.children[i].children[0]; /* 行の先頭は項目名 / each row starts with its label */
            rowLabels.push(rowLabel);
            labelWidth = Math.max(labelWidth, rowLabel.graphics.measureString(rowLabel.text)[0]);
        }
        for (var j = 0; j < rowLabels.length; j++) {
            rowLabels[j].preferredSize.width = Math.ceil(labelWidth) + 2;
        }
    }

    /**
     * URL を開くボタンを足す。URL が無ければディム表示にする。開いたらダイアログは閉じる
     * @param {Window} infoDialog - ボタンを置くダイアログ
     * @param {Group} buttonGroup - 足す先のグループ
     * @param {Object} labelEntry - ボタンの文言 { ja, en }
     * @param {string} targetUrl - 開く URL（空ならディム表示）
     * @returns {void}
     */
    function addVersionInfoLinkButton(infoDialog, buttonGroup, labelEntry, targetUrl) {
        var linkButton = buttonGroup.add("button", undefined, getVersionInfoLabel(labelEntry));
        if (!targetUrl) {
            linkButton.enabled = false;
            return;
        }
        linkButton.helpTip = targetUrl;
        linkButton.onClick = function () {
            openVersionInfoUrl(targetUrl);
            infoDialog.close();
        };
    }

    /**
     * 値をクリップボードへコピーするアイコンのボタンを足す。コピーしたらアイコンをチェックマークに変え、ポインタが外れたら戻す
     * @param {Group} infoRow - 足す先の行
     * @param {StaticText} valueLabel - コピーする値の statictext
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {Group} ボタンとして使う group
     */
    function addVersionInfoCopyButton(infoRow, valueLabel, isFromPalette) {
        var buttonSize = VERSION_INFO_COPY_BUTTON_SIZE;
        var copyButton = infoRow.add("group");
        copyButton.preferredSize = [buttonSize, buttonSize];
        copyButton.minimumSize = [buttonSize, buttonSize];
        copyButton.maximumSize = [buttonSize, buttonSize];
        copyButton.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.copyScriptName);
        copyButton.isPressed = false;
        copyButton.isCopied = false;

        copyButton.onDraw = function () {
            var buttonGraphics = copyButton.graphics;
            var fillColor = copyButton.isPressed ? VERSION_INFO_COPY_PRESSED_COLOR : VERSION_INFO_COPY_FILL_COLOR;
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0, 0, buttonSize, buttonSize);
            buttonGraphics.fillPath(buttonGraphics.newBrush(buttonGraphics.BrushType.SOLID_COLOR, fillColor));
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0.5, 0.5, buttonSize - 1, buttonSize - 1);
            buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_FRAME_COLOR, 1));
            if (copyButton.isCopied) drawVersionInfoCheckIcon(buttonGraphics, buttonSize);
            else drawVersionInfoCopyIcon(buttonGraphics, buttonSize);
        };

        /**
         * 状態を変えて描き直す（group には notify() が無いので隠して再表示する）
         * @param {boolean} isPressed - 押下中なら true
         * @param {boolean} isCopied - コピー済みなら true
         * @returns {void}
         */
        function repaint(isPressed, isCopied) {
            if (copyButton.isPressed === isPressed && copyButton.isCopied === isCopied) return;
            copyButton.isPressed = isPressed;
            copyButton.isCopied = isCopied;
            copyButton.hide();
            copyButton.show();
        }
        copyButton.addEventListener("mousedown", function () { repaint(true, false); });
        copyButton.addEventListener("mouseup", function () {
            if (!copyButton.isPressed) return;
            var hasCopied = copyVersionInfoText(valueLabel.text, isFromPalette);
            repaint(false, hasCopied);
            if (!hasCopied) alert(getVersionInfoLabel(VERSION_INFO_LABELS.alert.copyFailed));
        });
        /* 押したまま外へ出たとき・コピーのあと離れたときは元の見た目に戻す / Reset when the pointer leaves */
        copyButton.addEventListener("mouseout", function () { repaint(false, false); });
        return copyButton;
    }

    /**
     * コピーのアイコン（後ろの四角に前の四角を重ねた形）を描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCopyIcon(buttonGraphics, buttonSize) {
        var iconPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1);
        var sheetSize = 9;
        var offset = 3;
        var frontLeft = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var frontTop = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var backLeft = frontLeft - offset;
        var backTop = frontTop - offset;
        /* 後ろの四角は、前の四角に隠れる部分を描かない（上と左の辺だけ）/ Draw only the visible top and left of the back sheet */
        buttonGraphics.newPath();
        buttonGraphics.moveTo(frontLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, frontTop);
        buttonGraphics.strokePath(iconPen);
        buttonGraphics.newPath();
        buttonGraphics.rectPath(frontLeft, frontTop, sheetSize, sheetSize);
        buttonGraphics.strokePath(iconPen);
    }

    /**
     * コピーしたあとのチェックマークを描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCheckIcon(buttonGraphics, buttonSize) {
        var center = buttonSize / 2;
        buttonGraphics.newPath();
        buttonGraphics.moveTo(center - 4.5, center);
        buttonGraphics.lineTo(center - 1.5, center + 3);
        buttonGraphics.lineTo(center + 4.5, center - 3.5);
        buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1.5));
    }

    /**
     * 文字列をクリップボードにコピーする。
     * InDesign は AppleScript（Windows は VBScript）の doScript で書き込む。
     * Illustrator は文字列を直接送る API が無いので、一時レイヤーのテキストフレームを選んでコピーする（プレーンテキストも入る）。
     * Illustrator の常駐パレットからは DOM に触れないので、同じ処理を BridgeTalk でメインエンジンに渡す
     * @param {string} copyText - コピーする文字列
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {boolean} コピーの操作まで進めたら true
     */
    function copyVersionInfoText(copyText, isFromPalette) {
        try {
            if (VERSION_INFO_IS_INDESIGN) {
                if (File.fs === "Windows") {
                    app.doScript('CreateObject("htmlfile").parentWindow.clipboardData.setData "text", "' + copyText.replace(/"/g, '""') + '"', ScriptLanguage.VISUAL_BASIC);
                } else {
                    app.doScript('set the clipboard to "' + copyText.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"', ScriptLanguage.APPLESCRIPT_LANGUAGE);
                }
                return true;
            }
            var copyScript = buildVersionInfoCopyScript(copyText);
            if (!isFromPalette) return eval(copyScript) === true;
            var copyMessage = new BridgeTalk();
            copyMessage.target = BridgeTalk.appSpecifier; /* 複数バージョンを起動していても、今のバージョンに送る / the running version, even with several installed */
            copyMessage.body = copyScript;
            copyMessage.send();
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * Illustrator でコピーする処理を文字列で組み立てる（BridgeTalk でも送れるよう、関数の toString() は使わない）。
     * 一時レイヤーにテキストフレームを作り、フレームを選んでメニューコマンドでコピーして、レイヤーごと消す。
     * 文字として選択（textRange.select()）してコピーすると「オブジェクトをコピーできません」の警告が出るので使わない。
     * 選択と作業レイヤーは元に戻す。ドキュメントが無ければ一時ドキュメントを作って閉じる
     * @param {string} copyText - コピーする文字列
     * @returns {string} 実行すると true / false を返すスクリプト
     */
    function buildVersionInfoCopyScript(copyText) {
        return [
            "(function (copyText) {",
            "    var usingTempDoc = (app.documents.length === 0);",
            "    var targetDoc = usingTempDoc ? app.documents.add() : app.activeDocument;",
            "    var savedSelection = (!usingTempDoc && targetDoc.selection instanceof Array) ? targetDoc.selection : [];",
            "    var savedActiveLayer = usingTempDoc ? null : targetDoc.activeLayer;",
            "    var tempLayer = null;",
            "    var hasCopied = false;",
            "    try {",
            "        tempLayer = targetDoc.layers.add();",
            "        var tempFrame = tempLayer.textFrames.add();",
            "        tempFrame.contents = copyText;",
            "        app.redraw();",
            "        app.executeMenuCommand('deselectall');",
            "        tempFrame.selected = true;",
            "        app.redraw();",
            "        app.executeMenuCommand('copy');",
            "        app.redraw();",
            "        hasCopied = true;",
            "    } catch (e) { }",
            "    if (tempLayer) { try { tempLayer.remove(); } catch (e) { } }",
            "    if (usingTempDoc) { targetDoc.close(SaveOptions.DONOTSAVECHANGES); return hasCopied; }",
            "    try { targetDoc.activeLayer = savedActiveLayer; } catch (e) { }",
            "    targetDoc.selection = null;",
            "    for (var i = 0; i < savedSelection.length; i++) { try { savedSelection[i].selected = true; } catch (e) { } }",
            "    app.redraw();",
            "    return hasCopied;",
            "})(" + quoteVersionInfoString(copyText) + ");"
        ].join(" "); /* 行コメントを含まないので1行につなぐ（BridgeTalk の転送で改行が化けない）/ joined on one line; there are no line comments */
    }

    /**
     * 文字列を、スクリプトに埋め込める二重引用符のリテラルにする
     * @param {string} sourceText - 文字列
     * @returns {string} "…" の形のリテラル
     */
    function quoteVersionInfoString(sourceText) {
        return '"' + String(sourceText).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r/g, "\\r").replace(/\n/g, "\\n") + '"';
    }

    /**
     * ファイルを Finder で表示する。RevealInFinder.app があればファイルを選択した状態で表示し、無ければ入っているフォルダーを開く
     * @param {File} scriptFile - 表示するファイル
     * @returns {void}
     */
    function revealVersionInfoScriptFile(scriptFile) {
        /* .app は実体がフォルダーなので File.exists が false になる環境がある。Folder でも確かめる / An .app is a folder, so check with Folder too */
        var hasRevealApp = (File.fs === "Macintosh")
            && (new Folder(VERSION_INFO_REVEAL_APP_PATH).exists || new File(VERSION_INFO_REVEAL_APP_PATH).exists);
        /* アプリは一時ファイルからパスを読むので、書き出してから起動する / The app reads the path from a temp file */
        if (hasRevealApp && writeVersionInfoRevealPath(scriptFile.fsName) && new File(VERSION_INFO_REVEAL_APP_PATH).execute()) return;
        /* アプリが無い・起動できないときは、何も起きないままにせずフォルダーを開く / Otherwise open the enclosing folder */
        scriptFile.parent.execute();
    }

    /**
     * RevealInFinder.app に渡すパスを一時ファイルへ書き出す（UTF-8、改行なし）
     * @param {string} targetPath - 絶対パス
     * @returns {boolean} 書き出せたら true
     */
    function writeVersionInfoRevealPath(targetPath) {
        var pathFile = new File(VERSION_INFO_REVEAL_PATH_FILE);
        pathFile.encoding = "UTF-8";
        pathFile.lineFeed = "Unix";
        if (!pathFile.open("w")) return false;
        var hasWritten = pathFile.write(targetPath);
        pathFile.close();
        return hasWritten;
    }

    /**
     * URL を既定のブラウザーで開く（一時的なショートカットファイルを作って開く）
     * @param {string} targetUrl - 開く URL
     * @returns {void}
     */
    function openVersionInfoUrl(targetUrl) {
        var isMac = (File.fs === "Macintosh");
        var shortcutFile = new File(Folder.temp.fsName + "/VersionInfoLink" + (isMac ? ".webloc" : ".url"));
        shortcutFile.encoding = "UTF-8";
        shortcutFile.open("w");
        if (isMac) {
            var escapedUrl = targetUrl.replace(/&/g, "&amp;").replace(/</g, "&lt;");
            shortcutFile.write('<?xml version="1.0" encoding="UTF-8"?>\n'
                + '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n'
                + '<plist version="1.0"><dict><key>URL</key><string>' + escapedUrl + '</string></dict></plist>\n');
        } else {
            shortcutFile.write("[InternetShortcut]\r\nURL=" + targetUrl + "\r\n");
        }
        shortcutFile.close();
        shortcutFile.execute();
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

    /**
     * ［キャンセル］［OK］ボタンを追加する
     * @param {Group} buttonParent - 追加先のグループ
     * @returns {void}
     */
    function addCancelOkButtons(buttonParent) {
        buttonParent.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        buttonParent.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
    }

    /**
     * 対象範囲と除外ルールを選ぶダイアログを出す
     * @returns {object|null} { targetScope, exclusionOptions }。キャンセルなら null
     */
    function showScopeDialog() {
        var scopeDialog = new Window("dialog", getLabel(LABELS.dialog.scopeTitle));
        setupWindow(scopeDialog);

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

        var scopeButtonRow = addButtonRow(scopeDialog);
        addCancelOkButtons(scopeButtonRow.rightGroup);
        alignRightOnlyButtonRow(scopeButtonRow);

        addVersionFooter(scopeDialog);
        prepareDialogWindow(scopeDialog, SCRIPT_NAME);
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
        var previewDialog = new Window("dialog", getLabel(LABELS.dialog.previewTitle));
        setupWindow(previewDialog);

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

        var buttonRow = addButtonRow(previewDialog);
        var btnSelectAll = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.selectAll));
        btnSelectAll.helpTip = getLabel(LABELS.tooltip.selectAll);
        btnSelectAll.onClick = function () {
            checkAllRows(previewList);
        };
        addCancelOkButtons(buttonRow.rightGroup);
        alignRightOnlyButtonRow(buttonRow);

        prepareDialogWindow(previewDialog, SCRIPT_NAME + "_preview");
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
