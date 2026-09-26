#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のテキストに含まれる日付・曜日・連番・数値などを、一括して増減します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/IncrementDatesAndNumbers.md

### Overview

Increments or decrements the dates, weekday names, sequence numbers and other values found in the selected text, all at once.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/IncrementDatesAndNumbers.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IncrementDatesAndNumbers";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-11-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/IncrementDatesAndNumbers.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/IncrementDatesAndNumbers.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================
    var DIALOG_MARGINS        = 15;                /* ダイアログの余白 / dialog margins */
    var DIALOG_SPACING        = 10;                /* ダイアログ内の間隔 / dialog spacing */
    var PREVIEW_MARGINS       = [13, 10, 0, 10];   /* オリジナル／結果の行の余白 / margins of the preview rows */
    var PREVIEW_SPACING       = 6;                 /* オリジナル／結果の行どうしの間隔 / spacing between the preview rows */
    var PREVIEW_LABEL_SPACING = 6;                 /* 項目名と値の間隔 / gap between label and value */
    var STEP_PANEL_MARGINS    = [15, 20, 15, 13];  /* ［増減］パネルの余白 / margins of the step panel */
    var OPTION_PANEL_MARGINS  = [15, 20, 15, 10];  /* ［種別］［対象］パネルの余白 / margins of the type and target panels */
    var MODE_PANEL_SPACING    = 15;                /* ［種別］パネル内の間隔 / spacing in the type panel */
    var TARGET_PANEL_SPACING  = 2;                 /* ［対象］パネル内の間隔 / spacing in the target panel */
    var STEP_FIELD_CHARACTERS = 5;                 /* 値の入力欄の幅（文字数） / width of the step field (characters) */

    /* 結果欄の幅を測る見本。オリジナルが短いときもこの長さを確保する
       Sample that sizes the result field; used when the original text is short */
    var RESULT_WIDTH_SAMPLE     = "0000年00月00日";
    var RESULT_SAMPLE_MIN_CHARS = 8;

    // =========================================
    // 定数 / Constants
    // =========================================
    var DAY_NAMES = ["日", "月", "火", "水", "木", "金", "土"];
    var ENGLISH_DAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    var DAY_SYMBOLS = ["㊐", "㊊", "㊋", "㊌", "㊍", "㊎", "㊏"];
    var ERA_MAP = {
        "令和": 2019,
        "平成": 1989,
        "昭和": 1926,
        "大正": 1912,
        "明治": 1868
    };
    var CURRENT_YEAR = new Date().getFullYear();
    var LINE_BREAK = String.fromCharCode(13);

    // =========================================
    // 状態 / State
    // =========================================
    var stepValue = 1;       /* 1回の増減量 / amount per step */
    var shiftTarget = "day"; /* 増減する部分（"year" / "month" / "day"） / part to shift */

    /* ドット区切り2要素（例：12.1）の解釈（"number" または "date"）
       Mode for interpreting 2-part dot patterns like 12.1 ("number" or "date") */
    var dotPairMode = "number";

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILanguage();

    var LABELS = {
        dialog: {
            title: { ja: "値の増減", en: "Increment / Decrement values" }
        },
        panel: {
            step: { ja: "増減", en: "Increment" },
            mode: { ja: "種別", en: "Type" },
            target: { ja: "対象", en: "Target" }
        },
        fieldLabel: {
            original: { ja: "オリジナル", en: "Original" },
            result: { ja: "結果", en: "Result" },
            step: { ja: "値", en: "Value" }
        },
        radio: {
            modeNumber: { ja: "数字", en: "Number" },
            modeDate: { ja: "日付", en: "Date" }
        },
        tooltip: {
            step: {
                ja: "1回の増減で足す量です。負の値を入れると減らせます。",
                en: "Amount added per step. Enter a negative value to count down."
            },
            modeNumber: { ja: "テキストの中の数字を増減します。", en: "Increments the numbers in the text." },
            modeDate: {
                ja: "テキストを日付とみなして、年・月・日の単位で増減します。",
                en: "Reads the text as a date and steps it by year, month, or day."
            },
            target: { ja: "日付のどの部分を増減するかです。", en: "Which part of the date to step." }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            errorTitle: { ja: "エラー", en: "Error" },
            errorGeneric: { ja: "エラーが発生しました：", en: "An error occurred:" }
        }
    };

    /**
     * ドット区切りのパスで表示言語のラベルを取得する
     * @param {string} labelPath - "button.ok" のようなドット区切りキー
     * @returns {string} 表示言語のラベル。見つからない場合はパスをそのまま返す
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 選択の解析 / Selection analysis
    // =========================================

    /**
     * オブジェクトから最初のテキストフレームを探す（グループ内も再帰）
     * @param {PageItem} pageItem - 探索対象
     * @returns {TextFrame|null} 見つかったテキストフレーム
     */
    function findFirstTextFrame(pageItem) {
        if (!pageItem) return null;
        if (pageItem.typename === "TextFrame") return pageItem;
        if (pageItem.typename === "GroupItem") {
            var childItems = pageItem.pageItems;
            for (var i = 0; i < childItems.length; i++) {
                var foundFrame = findFirstTextFrame(childItems[i]);
                if (foundFrame) return foundFrame;
            }
        }
        return null;
    }

    /**
     * 選択から最初のテキストフレームを探す
     * @param {PageItem[]} currentSelection - ドキュメントの選択
     * @returns {TextFrame|null} 見つかったテキストフレーム
     */
    function findFirstTextFrameInSelection(currentSelection) {
        if (!currentSelection || currentSelection.length === 0) return null;
        for (var i = 0; i < currentSelection.length; i++) {
            var foundFrame = findFirstTextFrame(currentSelection[i]);
            if (foundFrame) return foundFrame;
        }
        return null;
    }

    /**
     * テキスト全体が年月日・ドット区切り・時刻のどれかなら、分割した各部分を返す
     * （［対象］パネルのラジオボタンの表示に使う）
     * @param {string} trimmedText - 前後の空白を除いたテキスト
     * @returns {{year: string, month: string, day: string, isDotPair: boolean}|null} 各部分。該当しなければ null
     */
    function splitDateParts(trimmedText) {
        /* 2025年11月21日のような和文年月日 / Japanese Y/M/D */
        var kanjiMatch = trimmedText.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日(?:[（(［\[]?(?:日|月|火|水|木|金|土)[）)\]］]?)?$/);
        if (kanjiMatch) return { year: kanjiMatch[1], month: kanjiMatch[2], day: kanjiMatch[3], isDotPair: false };

        /* 2025/11/21 や 2025/11/21（金）のようなスラッシュ区切り / Slash-separated date with optional weekday suffix */
        var slashMatch = trimmedText.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})(?:[ 　\t]*[（(［\[]?(?:日|月|火|水|木|金|土|Sun|Mon|Tue|Wed|Thu|Fri|Sat)[）)\]］]?)?$/);
        if (slashMatch) return { year: slashMatch[1], month: slashMatch[2], day: slashMatch[3], isDotPair: false };

        /* 29.3.2 / 29.4 などのドット区切り / dot-separated */
        var dotTripleMatch = trimmedText.match(/^(\d+)\.(\d+)\.(\d+)$/);
        if (dotTripleMatch) return { year: dotTripleMatch[1], month: dotTripleMatch[2], day: dotTripleMatch[3], isDotPair: false };

        /* 2要素のドット区切り（例：12.1）は「数字／日付」のどちらにもなり得るため、［種別］パネルを出す
           Two-part dot patterns (e.g. 12.1) can be numeric or date, so enable mode selection */
        var dotPairMatch = trimmedText.match(/^(\d+)\.(\d+)$/);
        if (dotPairMatch) return { year: dotPairMatch[1], month: dotPairMatch[2], day: "", isDotPair: true };

        /* 時刻 19:00 のようなパターンも分割対象とする（年=時、月=分として扱う）
           Treat time like 19:00 as a split target (map year->hour, month->minute) */
        var timeMatch = trimmedText.match(/^(\d{1,2}):(\d{2})$/);
        if (timeMatch) return { year: timeMatch[1], month: timeMatch[2], day: "", isDotPair: false };

        return null;
    }

    /**
     * 選択中の最初のテキストフレームから、プレビューの見本と［対象］パネルの内容を読み取る
     * @returns {{originalSample: string, hasYMDTarget: boolean, hasDot2Ambiguous: boolean, labelYear: string, labelMonth: string, labelDay: string}} 読み取り結果
     */
    function readSelectionSample() {
        var sampleInfo = {
            originalSample: "",
            hasYMDTarget: false,
            hasDot2Ambiguous: false,
            labelYear: "",
            labelMonth: "",
            labelDay: ""
        };

        /* ドキュメントが無いと activeDocument が例外になる / activeDocument throws without a document */
        try {
            var currentSelection = app.activeDocument.selection;
            if (!currentSelection || currentSelection.length === 0) return sampleInfo;
            var targetFrame = findFirstTextFrameInSelection(currentSelection);
            if (!targetFrame || targetFrame.typename !== "TextFrame") return sampleInfo;

            var trimmedText = targetFrame.contents.replace(/^\s+|\s+$/g, "");
            sampleInfo.originalSample = trimmedText.split(LINE_BREAK)[0];

            var dateParts = splitDateParts(trimmedText);
            if (dateParts) {
                sampleInfo.hasYMDTarget = true;
                sampleInfo.hasDot2Ambiguous = dateParts.isDotPair;
                sampleInfo.labelYear = dateParts.year;
                sampleInfo.labelMonth = dateParts.month;
                sampleInfo.labelDay = dateParts.day;
            }
        } catch (e) {}
        return sampleInfo;
    }

    // =========================================
    // 増減の計算 / Increment logic
    // =========================================

    /**
     * 数値を指定の桁数までゼロ埋めする
     * @param {number} numberValue - 数値
     * @param {number} digitCount - 桁数
     * @returns {string} ゼロ埋めした文字列
     */
    function pad(numberValue, digitCount) {
        var paddedText = String(numberValue);
        while (paddedText.length < digitCount) paddedText = "0" + paddedText;
        return paddedText;
    }

    /**
     * 数値文字列の整数部に3桁区切りのカンマを入れる
     * @param {string} numberText - 数値文字列
     * @returns {string} カンマ入りの文字列
     */
    function addCommas(numberText) {
        var numberParts = numberText.split(".");
        var intPart = numberParts[0];
        var decimalPart = numberParts.length > 1 ? "." + numberParts[1] : "";
        var isNegative = intPart.charAt(0) === "-";
        if (isNegative) intPart = intPart.substr(1);
        var withCommas = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
        return (isNegative ? "-" : "") + withCommas + decimalPart;
    }

    /**
     * 増減量を整数の段数にする（0 は 0、正は 1 以上、負は -1 以下）
     * @param {number} stepVal - 増減量
     * @returns {number} 整数の段数
     */
    function toWholeStep(stepVal) {
        if (stepVal === 0) return 0;
        if (stepVal > 0) return Math.max(1, Math.round(stepVal));
        return Math.min(-1, Math.round(stepVal));
    }

    /**
     * ［対象］の選択に応じて、日付の年・月・日のどれかを増減する
     * @param {Date} date - 対象の日付（直接書き換える）
     * @param {number} stepInt - 増減する段数
     * @returns {void}
     */
    function shiftDateByTarget(date, stepInt) {
        if (shiftTarget === "year") {
            date.setFullYear(date.getFullYear() + stepInt);
        } else if (shiftTarget === "month") {
            date.setMonth(date.getMonth() + stepInt);
        } else {
            date.setDate(date.getDate() + stepInt);
        }
    }

    /**
     * 「11月21日㊎」のような月日＋曜日記号を増減する
     * @param {string} text - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftMonthDayWithSymbol(text, stepInt) {
        var match = text.match(/(\d{1,2})月(\d{1,2})日(㊐|㊊|㊋|㊌|㊍|㊎|㊏)/);
        if (!match) return null;
        var shiftedDate = new Date(CURRENT_YEAR, parseInt(match[1], 10) - 1, parseInt(match[2], 10));
        shiftDateByTarget(shiftedDate, stepInt);
        var newText = (shiftedDate.getMonth() + 1) + "月" + shiftedDate.getDate() + "日" + DAY_SYMBOLS[shiftedDate.getDay()];
        return text.replace(match[0], newText);
    }

    /**
     * 「11月21日（金）」のような月日＋曜日を増減する
     * @param {string} text - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftMonthDayWithWeekday(text, stepInt) {
        var match = text.match(/(\d{1,2})月(\d{1,2})日([（([\[]?)(日|月|火|水|木|金|土)([）)\]]?)/);
        if (!match) return null;
        var shiftedDate = new Date(CURRENT_YEAR, parseInt(match[1], 10) - 1, parseInt(match[2], 10));
        shiftDateByTarget(shiftedDate, stepInt);
        var openBracket = match[3] || "";
        var closeBracket = match[5] || "";
        var newText = (shiftedDate.getMonth() + 1) + "月" + shiftedDate.getDate() + "日" + openBracket + DAY_NAMES[shiftedDate.getDay()] + closeBracket;
        return text.replace(match[0], newText);
    }

    /**
     * 「令和7年11月21日（金）」のような和暦の年月日＋曜日を増減する
     * @param {string} text - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftEraDateWithWeekday(text, stepInt) {
        var match = text.match(/(明治|大正|昭和|平成|令和)(\d{1,2})年(\d{1,2})月(\d{1,2})日([（([\[]?)(日|月|火|水|木|金|土)([）)\]]?)/);
        if (!match) return null;
        var baseYear = ERA_MAP[match[1]];
        var shiftedDate = new Date(baseYear + parseInt(match[2], 10) - 1, parseInt(match[3], 10) - 1, parseInt(match[4], 10));
        shiftDateByTarget(shiftedDate, stepInt);
        var newEra = match[1];
        for (var eraName in ERA_MAP) {
            if (shiftedDate.getFullYear() >= ERA_MAP[eraName]) newEra = eraName;
        }
        var eraYear = shiftedDate.getFullYear() - ERA_MAP[newEra] + 1;
        var openBracket = match[5] || "";
        var closeBracket = match[7] || "";
        var newText = newEra + eraYear + "年" + (shiftedDate.getMonth() + 1) + "月" + shiftedDate.getDate() + "日" + openBracket + DAY_NAMES[shiftedDate.getDay()] + closeBracket;
        return text.replace(match[0], newText);
    }

    /**
     * 「2025/11/21 (Fri)」「2025.11.21（金）」のような区切り付き日付＋曜日を増減する
     * @param {string} text - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftSeparatedDateWithWeekday(text, stepInt) {
        var match = text.match(/(\d{4})([\/.])(\d{1,2})\2(\d{1,2})(\s*)([（([\[]?)(日|月|火|水|木|金|土|Sun|Mon|Tue|Wed|Thu|Fri|Sat)([）)\]]?)/);
        if (!match) return null;
        var separator = match[2];
        var spacing = match[5] || "";
        var openBracket = match[6] || "";
        var weekdayText = match[7];
        var closeBracket = match[8] || "";
        var shiftedDate = new Date(parseInt(match[1], 10), parseInt(match[3], 10) - 1, parseInt(match[4], 10));
        shiftDateByTarget(shiftedDate, stepInt);
        var newDateText = shiftedDate.getFullYear() + separator + (shiftedDate.getMonth() + 1) + separator + shiftedDate.getDate();
        var newWeekday = /^[日月火水木金土]$/.test(weekdayText) ? DAY_NAMES[shiftedDate.getDay()] : ENGLISH_DAYS[shiftedDate.getDay()];
        var originalText = match[1] + separator + match[3] + separator + match[4] + spacing + openBracket + weekdayText + closeBracket;
        var replacedText = newDateText + spacing + openBracket + newWeekday + closeBracket;
        return text.replace(originalText, replacedText);
    }

    /* 曜日の無い数字だけの日付（20251121 / 2025/11/21 / 2025.11.21） / numeric dates without a weekday */
    var NUMERIC_DATE_FORMATS = [
        { regex: /(\d{4})(\d{2})(\d{2})/, format: "plain" },
        { regex: /(\d{4})\/(\d{1,2})\/(\d{1,2})/, format: "slash" },
        { regex: /(\d{4})\.(\d{1,2})\.(\d{1,2})/, format: "dot" }
    ];

    /**
     * 「20251121」「2025/11/21」「2025.11.21」のような数字だけの日付を増減する
     * @param {string} text - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftNumericDate(text, stepInt) {
        for (var i = 0; i < NUMERIC_DATE_FORMATS.length; i++) {
            var dateFormat = NUMERIC_DATE_FORMATS[i];
            var match = text.match(dateFormat.regex);
            if (!match) continue;
            var shiftedDate = new Date(parseInt(match[1], 10), parseInt(match[2], 10) - 1, parseInt(match[3], 10));
            shiftDateByTarget(shiftedDate, stepInt);
            var separator = (dateFormat.format === "slash") ? "/" : ".";
            var newText = (dateFormat.format === "plain") ?
                pad(shiftedDate.getFullYear(), 4) + pad(shiftedDate.getMonth() + 1, 2) + pad(shiftedDate.getDate(), 2) :
                shiftedDate.getFullYear() + separator + (shiftedDate.getMonth() + 1) + separator + shiftedDate.getDate();
            return text.replace(dateFormat.regex, newText);
        }
        return null;
    }

    /**
     * 「2025年11月20日(木)」のような曜日付きの年月日を増減する
     * @param {string} text - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftKanjiDateWithWeekday(text, stepInt) {
        var match = text.match(/(\d{4})年(\d{1,2})月(\d{1,2})日([（(［\[]?)(日|月|火|水|木|金|土)([）)\]]?)/);
        if (!match) return null;
        var openBracket = match[4] || "";
        var closeBracket = match[6] || "";
        var shiftedDate = new Date(parseInt(match[1], 10), parseInt(match[2], 10) - 1, parseInt(match[3], 10));
        shiftDateByTarget(shiftedDate, stepInt);
        var newText = shiftedDate.getFullYear() + "年" + (shiftedDate.getMonth() + 1) + "月" + shiftedDate.getDate() + "日" + openBracket + DAY_NAMES[shiftedDate.getDay()] + closeBracket;
        return text.replace(match[0], newText);
    }

    /**
     * 「2025年11月20日」のような曜日なしの年月日を増減する
     * （直後にかっこが続くものは曜日付きとして別に扱うので除く）
     * @param {string} text - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftKanjiDate(text, stepInt) {
        var match = text.match(/(\d{4})年(\d{1,2})月(\d{1,2})日(?![（(［\[])/);
        if (!match) return null;
        var shiftedDate = new Date(parseInt(match[1], 10), parseInt(match[2], 10) - 1, parseInt(match[3], 10));
        shiftDateByTarget(shiftedDate, stepInt);
        var newText = shiftedDate.getFullYear() + "年" + (shiftedDate.getMonth() + 1) + "月" + shiftedDate.getDate() + "日";
        return text.replace(match[0], newText);
    }

    /* 日付として増減するパターン（上から順に試し、最初に当たったものだけを使う）
       Date patterns, tried in order; the first match wins */
    var DATE_SHIFTERS = [
        shiftMonthDayWithSymbol,
        shiftMonthDayWithWeekday,
        shiftEraDateWithWeekday,
        shiftSeparatedDateWithWeekday,
        shiftNumericDate,
        shiftKanjiDateWithWeekday,
        shiftKanjiDate
    ];

    /**
     * 単独の曜日記号と曜日名を、増減の段数だけ送る
     * @param {string} text - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string} 置き換えたテキスト
     */
    function shiftWeekdayMarks(text, stepInt) {
        var weekdayShift = stepInt % 7;
        if (weekdayShift < 0) weekdayShift += 7;

        var updatedText = text;
        var symbolMatch = updatedText.match(/(㊐|㊊|㊋|㊌|㊍|㊎|㊏)/);
        if (symbolMatch) {
            var symbolIndex = DAY_SYMBOLS.indexOf(symbolMatch[1]);
            if (symbolIndex !== -1) {
                updatedText = updatedText.replace(symbolMatch[1], DAY_SYMBOLS[(symbolIndex + weekdayShift) % 7]);
            }
        }

        var weekdayMatch = updatedText.match(/([（([\[]?)(日|月|火|水|木|金|土)([）)\]]?)/);
        if (weekdayMatch) {
            var weekdayIndex = DAY_NAMES.indexOf(weekdayMatch[2]);
            if (weekdayIndex !== -1) {
                updatedText = updatedText.replace(weekdayMatch[0], (weekdayMatch[1] || "") + DAY_NAMES[(weekdayIndex + weekdayShift) % 7] + (weekdayMatch[3] || ""));
            }
        }
        return updatedText;
    }

    /**
     * 1900〜2099 の4桁の年を増減する
     * @param {string} text - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftYearNumber(text, stepInt) {
        var yearMatch = text.match(/\b(19|20)\d{2}\b/);
        if (!yearMatch) return null;
        var yearText = yearMatch[0];
        return text.replace(yearText, String(parseInt(yearText, 10) + stepInt));
    }

    /**
     * 「29.3.2」「12.1」のようなドット区切りの数字を増減する
     * @param {string} text - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftDotNumbers(text, stepInt) {
        var dotTripleMatch = text.match(/^(\d+)\.(\d+)\.(\d+)$/);
        if (dotTripleMatch) {
            /* 3要素のドット区切りは数値として扱う / Three-part dot patterns stay numeric */
            var firstPart = parseInt(dotTripleMatch[1], 10);
            var secondPart = parseInt(dotTripleMatch[2], 10);
            var thirdPart = parseInt(dotTripleMatch[3], 10);
            if (shiftTarget === "year") firstPart += stepInt;
            else if (shiftTarget === "month") secondPart += stepInt;
            else if (shiftTarget === "day") thirdPart += stepInt;
            return firstPart + "." + secondPart + "." + thirdPart;
        }

        var dotPairMatch = text.match(/^(\d+)\.(\d+)$/);
        if (!dotPairMatch) return null;
        return (dotPairMode === "date") ? shiftDotPairAsDate(dotPairMatch, stepInt) : shiftDotPairAsNumber(dotPairMatch, stepInt);
    }

    /**
     * 2要素のドット区切りを「月.日」とみなして増減する（年は今年で仮置き）
     * ［対象］が前半なら月、それ以外なら日を増減し、元の桁数でゼロ埋めする（例：12.05 → 12.06）
     * @param {string[]} dotPairMatch - /^(\d+)\.(\d+)$/ の一致結果
     * @param {number} stepInt - 増減する段数
     * @returns {string} 置き換えたテキスト
     */
    function shiftDotPairAsDate(dotPairMatch, stepInt) {
        var monthWidth = dotPairMatch[1].length;
        var dayWidth = dotPairMatch[2].length;
        var shiftedDate = new Date(CURRENT_YEAR, parseInt(dotPairMatch[1], 10) - 1, parseInt(dotPairMatch[2], 10));

        if (shiftTarget === "year") {
            /* 「年」ターゲット（前半）は月を増減 / the first part shifts the month */
            shiftedDate.setMonth(shiftedDate.getMonth() + stepInt);
        } else {
            /* それ以外は日を増減 / other targets shift the day */
            shiftedDate.setDate(shiftedDate.getDate() + stepInt);
        }

        var newMonth = shiftedDate.getMonth() + 1;
        var newDay = shiftedDate.getDate();
        var newMonthText = (monthWidth > 1) ? pad(newMonth, monthWidth) : String(newMonth);
        var newDayText = (dayWidth > 1) ? pad(newDay, dayWidth) : String(newDay);
        return newMonthText + "." + newDayText;
    }

    /**
     * 2要素のドット区切りを数字として増減する
     * ［対象］が前半なら前半だけ、後半なら後半を増減して桁あふれを前半へ繰り上げ／繰り下げる
     * @param {string[]} dotPairMatch - /^(\d+)\.(\d+)$/ の一致結果
     * @param {number} stepInt - 増減する段数
     * @returns {string} 置き換えたテキスト
     */
    function shiftDotPairAsNumber(dotPairMatch, stepInt) {
        var firstPart = parseInt(dotPairMatch[1], 10);
        var secondPart = parseInt(dotPairMatch[2], 10);

        if (shiftTarget === "year") {
            firstPart += stepInt;
        } else {
            /* 後半の桁数から基数を決め、合計を分解し直す / base from the digit count, then split the total back */
            var base = Math.pow(10, String(Math.abs(secondPart)).length || 1);
            var total = firstPart * base + secondPart + stepInt;
            var newFirstPart = (total >= 0) ? Math.floor(total / base) : -Math.ceil(-total / base);
            secondPart = total - newFirstPart * base;
            firstPart = newFirstPart;
        }
        return firstPart + "." + secondPart;
    }

    /**
     * 「19:00」のような時刻を増減する（［対象］が前半なら時、それ以外は分。時は24時間でループ）
     * @param {string} text - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftTime(text, stepInt) {
        var timeMatch = text.match(/(\d{1,2}):(\d{2})/);
        if (!timeMatch) return null;
        var hour = parseInt(timeMatch[1], 10);
        var minute = parseInt(timeMatch[2], 10);

        if (shiftTarget === "year") {
            hour += stepInt;
        } else {
            minute += stepInt;
        }

        /* 分のあふれを時へ繰り上げ／繰り下げる / carry minute overflow into the hour */
        while (minute < 0) {
            minute += 60;
            hour -= 1;
        }
        while (minute >= 60) {
            minute -= 60;
            hour += 1;
        }
        hour = ((hour % 24) + 24) % 24;

        var newTime = String(hour) + ":" + (minute < 10 ? "0" + String(minute) : String(minute));
        return text.replace(timeMatch[0], newTime);
    }

    /**
     * 最初に現れる数値（カンマ区切り・小数を含む）を増減する。小数は最下位の桁を1単位とする
     * @param {string} text - 対象のテキスト
     * @param {number} stepVal - 増減量
     * @returns {string} 置き換えたテキスト（数値が無ければそのまま）
     */
    function shiftFirstNumber(text, stepVal) {
        var numberMatch = text.match(/-?\d+(?:,\d{3})*(?:\.\d+)?/);
        if (!numberMatch) return text;
        var numberText = numberMatch[0];
        var rawNumberText = numberText.replace(/,/g, "");
        var decimalMatch = rawNumberText.match(/\.(\d+)/);
        var decimalDigits = decimalMatch ? decimalMatch[1].length : 0;
        var unitStep = decimalDigits > 0 ? 1 / Math.pow(10, decimalDigits) : 1;
        var shiftedValue = parseFloat(rawNumberText) + unitStep * stepVal;
        return text.replace(numberText, addCommas(shiftedValue.toFixed(decimalDigits)));
    }

    /**
     * 1行分のテキストに含まれる日付・曜日・時刻・数値を、現在の設定で増減する
     * @param {string} text - 対象の1行
     * @returns {string} 増減後のテキスト（失敗したら元のまま）
     */
    function updateLine(text) {
        /* 途中で例外が出たら元の行を返す（ExtendScript の配列には indexOf が無いことがある）
           Return the line untouched on any error (ExtendScript arrays may lack indexOf) */
        try {
            var updatedText = text.replace(/^\s+|\s+$/g, "");
            var stepVal = (typeof stepValue === "number" && !isNaN(stepValue)) ? stepValue : 1;
            var stepInt = toWholeStep(stepVal);

            for (var i = 0; i < DATE_SHIFTERS.length; i++) {
                var shiftedDateText = DATE_SHIFTERS[i](updatedText, stepInt);
                if (shiftedDateText !== null) return shiftedDateText;
            }

            updatedText = shiftWeekdayMarks(updatedText, stepInt);

            var shiftedText = shiftYearNumber(updatedText, stepInt);
            if (shiftedText === null) shiftedText = shiftDotNumbers(updatedText, stepInt);
            if (shiftedText === null) shiftedText = shiftTime(updatedText, stepInt);
            if (shiftedText !== null) return shiftedText;

            return shiftFirstNumber(updatedText, stepVal);
        } catch (e) {
            $.writeln("Error in updateLine: " + e.message);
            return text;
        }
    }

    // =========================================
    // 適用 / Apply
    // =========================================

    /**
     * テキストフレームの各行を増減する（グループ内も再帰）
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {void}
     */
    function updatePageItem(pageItem) {
        if (!pageItem) return;
        if (pageItem.typename === "TextFrame") {
            var lines = pageItem.contents.split(LINE_BREAK);
            var updatedLines = [];
            for (var i = 0; i < lines.length; i++) {
                updatedLines.push(updateLine(lines[i]));
            }
            pageItem.contents = updatedLines.join(LINE_BREAK);
            return;
        }
        if (pageItem.typename === "GroupItem") {
            var childItems = pageItem.pageItems;
            for (var j = 0; j < childItems.length; j++) {
                updatePageItem(childItems[j]);
            }
        }
    }

    /**
     * 選択中のオブジェクトに増減を適用する
     * @returns {void}
     */
    function applyToSelection() {
        try {
            var selectedItems = app.activeDocument.selection;
            for (var i = 0; i < selectedItems.length; i++) {
                updatePageItem(selectedItems[i]);
            }
        } catch (e) {
            alert(getLabel("alert.errorGeneric") + LINE_BREAK + e.message, getLabel("alert.errorTitle"));
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 結果欄の右に出す数値を求める（純粋な数値どうしなら重複するので出さない）
     * @param {string} originalSample - オリジナルの1行
     * @param {string} previewText - 増減後の1行
     * @returns {number|string} 表示する数値。出さないときは空文字
     */
    function computeCalcValue(originalSample, previewText) {
        var pureNumberPattern = /^\s*-?\d+(?:,\d{3})*(?:\.\d+)?\s*$/;
        if (pureNumberPattern.test(originalSample) && pureNumberPattern.test(previewText)) return "";

        var numberPattern = /-?\d+(?:,\d{3})*(?:\.\d+)?/;
        var originalMatch = originalSample.match(numberPattern);
        var resultMatch = previewText.match(numberPattern);
        if (!originalMatch || !resultMatch) return "";

        var originalNumber = Number(originalMatch[0].replace(/,/g, ""));
        var resultNumber = Number(resultMatch[0].replace(/,/g, ""));
        return (!isNaN(originalNumber) && !isNaN(resultNumber)) ? resultNumber : "";
    }

    /**
     * オリジナル／結果の2行を追加する
     * @param {Window} incrementDialog - 追加先のダイアログ
     * @param {string} originalSample - オリジナルの1行
     * @returns {{resultValueText: StaticText, calcValueText: StaticText}} 更新する表示欄
     */
    function addPreviewRows(incrementDialog, originalSample) {
        var previewGroup = incrementDialog.add("group");
        previewGroup.orientation = "column";
        previewGroup.alignChildren = ["left", "top"];
        previewGroup.margins = PREVIEW_MARGINS;
        previewGroup.spacing = PREVIEW_SPACING;

        var originalRow = previewGroup.add("group");
        originalRow.orientation = "row";
        originalRow.alignChildren = ["left", "center"];
        originalRow.spacing = PREVIEW_LABEL_SPACING;
        var originalLabel = originalRow.add("statictext", undefined, labelText("fieldLabel.original"));
        originalRow.add("statictext", undefined, originalSample || "");

        var resultRow = previewGroup.add("group");
        resultRow.orientation = "row";
        resultRow.alignChildren = ["left", "center"];
        resultRow.spacing = PREVIEW_LABEL_SPACING;
        var resultLabel = resultRow.add("statictext", undefined, labelText("fieldLabel.result"));
        var resultValueText = resultRow.add("statictext", undefined, "");
        var calcValueText = resultRow.add("statictext", undefined, "");

        /* 項目名の幅をそろえて右揃えにし、結果欄の幅を確保する / align the labels and reserve the result width */
        try {
            var dialogGraphics = incrementDialog.graphics;
            var labelWidth = Math.ceil(Math.max(
                dialogGraphics.measureString(originalLabel.text)[0],
                dialogGraphics.measureString(resultLabel.text)[0]
            ));
            originalLabel.preferredSize = [labelWidth, -1];
            resultLabel.preferredSize = [labelWidth, -1];
            originalLabel.justify = "right";
            resultLabel.justify = "right";

            /* 「12.1」など短いオリジナルは、日付相当の見本で幅を確保（結果の切れを防ぐ）
               Short originals fall back to a date-length sample to avoid truncating the result */
            var widthSample = originalSample || RESULT_WIDTH_SAMPLE;
            if (widthSample.length < RESULT_SAMPLE_MIN_CHARS) widthSample = RESULT_WIDTH_SAMPLE;
            resultValueText.preferredSize = [Math.ceil(dialogGraphics.measureString(widthSample)[0]), -1];
        } catch (e) {}

        return { resultValueText: resultValueText, calcValueText: calcValueText };
    }

    /**
     * ［種別］パネル（数字／日付）を追加する。「12.1」のような2要素のドット区切りのときだけ使う
     * @param {Window} incrementDialog - 追加先のダイアログ
     * @param {Function} onModeChange - 切り替えたときに呼ぶ関数
     * @returns {void}
     */
    function addModePanel(incrementDialog, onModeChange) {
        var modePanel = incrementDialog.add("panel", undefined, getLabel("panel.mode"));
        modePanel.orientation = "row";
        modePanel.alignChildren = ["left", "center"];
        modePanel.margins = OPTION_PANEL_MARGINS;
        modePanel.spacing = MODE_PANEL_SPACING;

        var numberModeRadio = modePanel.add("radiobutton", undefined, getLabel("radio.modeNumber"));
        numberModeRadio.helpTip = getLabel("tooltip.modeNumber");
        var dateModeRadio = modePanel.add("radiobutton", undefined, getLabel("radio.modeDate"));
        dateModeRadio.helpTip = getLabel("tooltip.modeDate");

        /* 既定は「数字」 / default is Number */
        dotPairMode = "number";
        numberModeRadio.value = true;

        numberModeRadio.onClick = function () {
            dotPairMode = "number";
            onModeChange();
        };
        dateModeRadio.onClick = function () {
            dotPairMode = "date";
            onModeChange();
        };
    }

    /**
     * ［対象］パネルを追加する（分割できた部分ごとのラジオボタン。既定は最後の部分）
     * @param {Window} incrementDialog - 追加先のダイアログ
     * @param {{labelYear: string, labelMonth: string, labelDay: string}} sampleInfo - readSelectionSample() の結果
     * @returns {{yearRadio: (RadioButton|undefined), monthRadio: (RadioButton|undefined), dayRadio: (RadioButton|undefined)}} 作ったラジオボタン
     */
    function addTargetPanel(incrementDialog, sampleInfo) {
        var targetPanel = incrementDialog.add("panel", undefined, getLabel("panel.target"));
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["left", "top"];
        targetPanel.margins = OPTION_PANEL_MARGINS;
        targetPanel.spacing = TARGET_PANEL_SPACING;

        var targetRadios = {};
        if (sampleInfo.labelYear !== "") targetRadios.yearRadio = targetPanel.add("radiobutton", undefined, sampleInfo.labelYear);
        if (sampleInfo.labelMonth !== "") targetRadios.monthRadio = targetPanel.add("radiobutton", undefined, sampleInfo.labelMonth);
        if (sampleInfo.labelDay !== "") targetRadios.dayRadio = targetPanel.add("radiobutton", undefined, sampleInfo.labelDay);

        if (targetRadios.yearRadio) targetRadios.yearRadio.helpTip = getLabel("tooltip.target");
        if (targetRadios.monthRadio) targetRadios.monthRadio.helpTip = getLabel("tooltip.target");
        if (targetRadios.dayRadio) targetRadios.dayRadio.helpTip = getLabel("tooltip.target");

        if (targetRadios.dayRadio) targetRadios.dayRadio.value = true;
        else if (targetRadios.monthRadio) targetRadios.monthRadio.value = true;
        else if (targetRadios.yearRadio) targetRadios.yearRadio.value = true;

        return targetRadios;
    }

    /**
     * ［対象］のラジオボタンから増減する部分を読む
     * @param {Object|null} targetRadios - addTargetPanel() の結果（パネルが無ければ null）
     * @returns {string} "year" / "month" / "day"
     */
    function readShiftTarget(targetRadios) {
        if (!targetRadios) return "day";
        if (targetRadios.yearRadio && targetRadios.yearRadio.value) return "year";
        if (targetRadios.monthRadio && targetRadios.monthRadio.value) return "month";
        return "day";
    }

    /**
     * ダイアログを表示し、OK で選択中のテキストに増減を適用する
     * @returns {void}
     */
    function showIncrementDialog() {
        var sampleInfo = readSelectionSample();
        var originalSample = sampleInfo.originalSample;

        var incrementDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        incrementDialog.orientation = "column";
        incrementDialog.alignChildren = ["fill", "top"];
        incrementDialog.margins = DIALOG_MARGINS;
        incrementDialog.spacing = DIALOG_SPACING;

        var previewTexts = addPreviewRows(incrementDialog, originalSample);

        /**
         * 現在の設定でオリジナルの1行を変換し、結果欄に出す
         * @returns {void}
         */
        function updateResultPreview() {
            if (!originalSample) {
                previewTexts.resultValueText.text = "";
                previewTexts.calcValueText.text = "";
                return;
            }
            var previewText = updateLine(originalSample);
            previewTexts.resultValueText.text = previewText;
            /* 年月日などを分割できるときは右側の数値を出さない / no extra number for split targets */
            previewTexts.calcValueText.text = sampleInfo.hasYMDTarget ? "" : computeCalcValue(originalSample, previewText);
        }

        var stepPanel = incrementDialog.add("panel", undefined, getLabel("panel.step"));
        stepPanel.orientation = "column";
        stepPanel.alignChildren = ["left", "top"];
        stepPanel.margins = STEP_PANEL_MARGINS;
        stepPanel.spacing = 10;

        var stepRow = stepPanel.add("group");
        stepRow.orientation = "row";
        stepRow.alignChildren = ["left", "center"];
        stepRow.spacing = 5;
        stepRow.add("statictext", undefined, getLabel("fieldLabel.step"));
        var stepInput = stepRow.add("edittext", undefined, "1");
        stepInput.helpTip = getLabel("tooltip.step");
        stepInput.characters = STEP_FIELD_CHARACTERS;
        changeValueByArrowKey(stepInput);

        stepInput.onChanging = function () {
            var parsedStep = parseInt(stepInput.text, 10);
            if (isNaN(parsedStep)) return;
            stepValue = parsedStep;
            stepInput.text = String(parsedStep);
            updateResultPreview();
        };
        stepInput.onChange = stepInput.onChanging;

        incrementDialog.onShow = function () {
            stepInput.active = true;
            updateResultPreview();
        };

        if (sampleInfo.hasDot2Ambiguous) {
            addModePanel(incrementDialog, updateResultPreview);
        }

        var targetRadios = sampleInfo.hasYMDTarget ? addTargetPanel(incrementDialog, sampleInfo) : null;

        /**
         * ［対象］の選択を反映し、必要なら増減値を 1 に戻して入力欄を編集状態にする
         * @param {boolean} resetStep - 増減値を 1 に戻すなら true
         * @returns {void}
         */
        function applyTargetFromRadios(resetStep) {
            shiftTarget = readShiftTarget(targetRadios);
            if (resetStep) {
                stepValue = 1;
                stepInput.text = "1";
                stepInput.active = true;
            }
            updateResultPreview();
        }

        if (targetRadios) {
            var onTargetClick = function () {
                applyTargetFromRadios(true);
            };
            if (targetRadios.yearRadio) targetRadios.yearRadio.onClick = onTargetClick;
            if (targetRadios.monthRadio) targetRadios.monthRadio.onClick = onTargetClick;
            if (targetRadios.dayRadio) targetRadios.dayRadio.onClick = onTargetClick;

            /* 既定の選択を反映（ここでは増減値を戻さない） / sync the default selection without resetting the step */
            applyTargetFromRadios(false);
        }

        var btnRowGroup = incrementDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["right", "center"];
        btnRowGroup.alignment = ["fill", "bottom"];
        btnRowGroup.spacing = 10;

        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"));
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"));

        btnOK.onClick = function () {
            var parsedStep = parseInt(stepInput.text, 10);
            if (isNaN(parsedStep)) parsedStep = 1;
            stepValue = parsedStep;
            stepInput.text = String(parsedStep);

            /* ラジオボタンの状態を最終的な増減対象に反映（増減値は維持） / sync the target, keeping the step */
            applyTargetFromRadios(false);

            applyToSelection();
            incrementDialog.close(1);
        };

        btnCancel.onClick = function () {
            incrementDialog.close(0);
        };

        incrementDialog.layout.layout(true);
        incrementDialog.center();
        incrementDialog.show();
    }

    /**
     * 数値欄で ↑↓ キーによる増減を有効にする（Shift で 10 刻み、Option で 0.1 刻み）
     * @param {EditText} editText - 対象の入力欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        if (!editText) return;
        editText.addEventListener("keydown", function(event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;
            var keyboardState = ScriptUI.environment.keyboardState;
            var delta = 1;
            if (keyboardState.shiftKey) {
                delta = 10;
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    event.preventDefault();
                }
            } else if (keyboardState.altKey) {
                delta = 0.1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            }
            if (keyboardState.altKey) {
                value = Math.round(value * 10) / 10;
            } else {
                value = Math.round(value);
            }
            editText.text = value;
            if (typeof editText.onChange === 'function') {
                try {
                    editText.onChange();
                } catch (e) {}
            }
        });
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    showIncrementDialog();

})();
