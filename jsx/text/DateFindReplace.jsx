#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内のテキストフレームから日付を検索し、選択した項目だけを置換します。
オブジェクトが選択されている場合は、その選択範囲のテキストフレームだけを検索対象にします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DateFindReplace.md

### Overview

Finds dates in the text frames of the document and replaces only the ones you tick.
When objects are selected, only the text frames within that selection are searched.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DateFindReplace.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DateFindReplace";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DateFindReplace.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DateFindReplace.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10];  /* パネルの余白 [左,上,右,下] / Panel margins */
    var PANEL_SPACING = 6;                 /* パネル内の間隔 / Spacing inside panels */
    var FOUND_PANEL_MIN_WIDTH = 240;       /* ［見つかった日付］パネルの最小幅 / Minimum width of the found-dates panel */
    var DATE_ROW_INDENT = 20;              /* 日付行の左インデント / Left indent of the date rows */
    var DATE_CHECKBOX_WIDTH = 220;         /* 日付チェックボックスの幅（ラベル切れ対策）/ Width of the date checkboxes */
    var CHECK_LABEL_WIDTH = 130;           /* 確認パネルの項目名の幅（値の左端をそろえる）/ Label width in the check panel */

    // =========================================
    // 和暦 / Japanese eras
    // =========================================

    /* 令和元年 = 2019 年（令和Y = 西暦 Y + 2018）／平成元年 = 1989 年（平成Y = 西暦 Y + 1988） */
    var REIWA_BASE_YEAR = 2018;
    var HEISEI_BASE_YEAR = 1988;

    /* 各元号の有効範囲（getValidationErrorMessage で使用） */
    var REIWA_START_DATE = new Date(2019, 4, 1);                 /* 2019/5/1 〜 */
    var HEISEI_START_DATE = new Date(1989, 0, 8);                /* 1989/1/8 〜 */
    var HEISEI_END_DATE_EXCLUSIVE = new Date(2019, 4, 1);        /* 〜 2019/4/30 */

    /* 確認パネルの和暦表示に使う元号表（新しい順）/ Era table for the check panel (newest first) */
    var ERA_TABLE = [
        { name: "令和", startDate: REIWA_START_DATE, baseYear: REIWA_BASE_YEAR },   /* 2019/5/1〜 */
        { name: "平成", startDate: HEISEI_START_DATE, baseYear: HEISEI_BASE_YEAR }, /* 1989/1/8〜2019/4/30 */
        { name: "昭和", startDate: new Date(1926, 11, 25), baseYear: 1925 },        /* 1926/12/25〜1989/1/7 */
        { name: "大正", startDate: new Date(1912, 6, 30), baseYear: 1911 },         /* 1912/7/30〜1926/12/24 */
        { name: "明治", startDate: new Date(1868, 8, 8), baseYear: 1867 }           /* 1868/9/8〜1912/7/29 */
    ];

    /**
     * 令和の年を西暦にする
     * @param {number} reiwaYear - 令和の年
     * @returns {number} 西暦
     */
    function reiwaToGregorian(reiwaYear) { return reiwaYear + REIWA_BASE_YEAR; }

    /**
     * 西暦を令和の年にする
     * @param {number} gregorianYear - 西暦
     * @returns {number} 令和の年
     */
    function gregorianToReiwa(gregorianYear) { return gregorianYear - REIWA_BASE_YEAR; }

    /**
     * 平成の年を西暦にする
     * @param {number} heiseiYear - 平成の年
     * @returns {number} 西暦
     */
    function heiseiToGregorian(heiseiYear) { return heiseiYear + HEISEI_BASE_YEAR; }

    /**
     * 西暦を平成の年にする
     * @param {number} gregorianYear - 西暦
     * @returns {number} 平成の年
     */
    function gregorianToHeisei(gregorianYear) { return gregorianYear - HEISEI_BASE_YEAR; }

    // =========================================
    // 選択肢 / Choices
    // =========================================

    /*
       出力フォーマット選択。
       「元の形式を保持」を選ぶと、各マッチの元形式（区切り文字・元号）を維持して置換する。
       曜日表記は「元の形式を保持」でも隣の「曜日」ドロップダウンの選択を反映する。
       それ以外は、選択した形式で全マッチを統一して書き換える。
    */
    var FORMAT_VALUES = ["preserve", "jp", "jp-md", "dot", "dot-md", "slash", "slash-md", "reiwa-jp", "r-dot", "r-slash", "heisei-jp", "h-dot", "h-slash"];
    var FORMAT_LABELS = [
        "元の形式を保持",
        "YYYY年M月D日",
        "M月D日",
        "YYYY.M.D",
        "M.D",
        "YYYY/M/D",
        "M/D",
        "令和Y年M月D日",
        "RY.M.D",
        "RY/M/D",
        "平成Y年M月D日",
        "HY.M.D",
        "HY/M/D"
    ];

    /* 曜日サフィックスのスタイル（「元の形式を保持」以外の出力に付与） */
    var WEEKDAY_VALUES = ["none", "kanji", "medium", "long", "full-paren", "half-paren", "en-short", "en-full"];
    var WEEKDAY_LABELS = ["なし", "火", "火曜", "火曜日", "（火）", "(火)", "Tue", "Tuesday"];

    /* 曜日生成用の文字テーブル */
    var WEEKDAY_KANJI = ["日", "月", "火", "水", "木", "金", "土"];
    var WEEKDAY_EN_SHORT = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    var WEEKDAY_EN_FULL = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

    /* ↑↓キーの tooltip / Tooltip for the arrow keys */
    var ARROW_KEY_HELP = "↑↓キーで増減（Shift：±10）";

    // =========================================
    // 共通ヘルパー / Shared helpers
    // =========================================

    /**
     * エラー件数のメッセージを作る
     * @param {number} errorCount - 処理できなかった件数
     * @returns {string} メッセージ
     */
    function formatErrorMessage(errorCount) {
        return errorCount + "件はロック等の理由で処理できませんでした。";
    }

    /**
     * パネルの共通設定
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = "left";
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            targetPanel.spacing = spacing;
        }
    }

    /**
     * テキストフレームの中心が乗っているアートボードの番号を返す
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @returns {number} アートボードの番号（どれにも乗っていなければ -1）
     */
    function getArtboardIndex(doc, targetItem) {
        var itemBounds = targetItem.geometricBounds;
        var centerX = (itemBounds[0] + itemBounds[2]) / 2;
        var centerY = (itemBounds[1] + itemBounds[3]) / 2;

        for (var artboardIdx = 0; artboardIdx < doc.artboards.length; artboardIdx++) {
            var artboardRect = doc.artboards[artboardIdx].artboardRect;
            var minX = Math.min(artboardRect[0], artboardRect[2]);
            var maxX = Math.max(artboardRect[0], artboardRect[2]);
            var minY = Math.min(artboardRect[1], artboardRect[3]);
            var maxY = Math.max(artboardRect[1], artboardRect[3]);
            if (centerX >= minX && centerX <= maxX && centerY >= minY && centerY <= maxY) {
                return artboardIdx;
            }
        }
        return -1;
    }

    /**
     * ↑↓キーで値を増減する（Shiftで±10、1未満にはしない）
     * @param {EditText} editText - 対象の入力欄
     * @param {Function} [onValueChanged] - 値を変えたあとに呼ぶ処理
     * @returns {void}
     */
    function changeValueByArrowKey(editText, onValueChanged) {
        editText.addEventListener("keydown", function (event) {
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
                    if (value < 1) value = 1;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    if (value < 1) value = 1;
                    event.preventDefault();
                }
            }
            value = Math.round(value);
            if (value < 1) value = 1;
            editText.text = value;
            if (typeof onValueChanged === "function") {
                onValueChanged();
            }
        });
    }

    /**
     * dropdownlist の選択値を返す（選択が外れているときは代わりの値）
     * @param {DropDownList} dropdown - 対象のドロップダウン
     * @param {string[]} valuesArray - 項目ごとの内部値
     * @param {string} fallback - 選択が無いときの値
     * @returns {string} 選択中の内部値
     */
    function getDropdownValue(dropdown, valuesArray, fallback) {
        if (dropdown.selection !== null && dropdown.selection !== undefined) {
            var selectedIndex = dropdown.selection.index;
            if (typeof selectedIndex === 'number' && selectedIndex >= 0 && selectedIndex < valuesArray.length) {
                return valuesArray[selectedIndex];
            }
        }
        return fallback;
    }

    // =========================================
    // 日付の解析 / Parsing dates
    // =========================================

    /**
     * 曜日の文字列の形式名を返す（WEEKDAY_VALUES と同じ命名）
     * @param {string} weekdayText - 判定する文字列
     * @param {boolean} allowSingleKanji - 漢字1文字（「金」など）も曜日として認めるか
     * @returns {string|null} 形式名（曜日でなければ null）
     */
    function classifyWeekdayText(weekdayText, allowSingleKanji) {
        if (/^[日月火水木金土]曜日$/.test(weekdayText)) return 'long';           /* 例：金曜日 */
        if (/^[日月火水木金土]曜$/.test(weekdayText)) return 'medium';           /* 例：金曜 */
        if (/^\([日月火水木金土]\)$/.test(weekdayText)) return 'half-paren';   /* 例：(金) */
        if (/^（[日月火水木金土]）$/.test(weekdayText)) return 'full-paren';    /* 例：（金） */
        if (allowSingleKanji && /^[日月火水木金土]$/.test(weekdayText)) return 'kanji';
        if (/^(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat)day$/.test(weekdayText)) return 'en-full';
        if (/^(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat)$/.test(weekdayText)) return 'en-short';
        return null;
    }

    /**
     * 解析済みの日付に付いている曜日サフィックスの形式を返す（漢字1文字は対象外）
     * @param {Object} parsedDate - detectFormatAndParse() の戻り値
     * @returns {string|null} 形式名（サフィックスが無い・曜日でなければ null）
     */
    function getWeekdaySuffixStyle(parsedDate) {
        if (!parsedDate || !parsedDate.parts || !parsedDate.parts.suffix) return null;
        return classifyWeekdayText(parsedDate.parts.suffix, false);
    }

    /**
     * フレーム全体が曜日だけかを判定し、形式名を返す（単独漢字曜日はここでのみ許可）
     * @param {string} frameContent - テキストフレームの内容
     * @returns {string|null} 形式名（曜日だけでなければ null）
     */
    function detectWeekdayOnlyFrame(frameContent) {
        return classifyWeekdayText(String(frameContent).replace(/^\s+|\s+$/g, ""), true);
    }

    /**
     * 元号年の文字列化。漢字フォーマット（reiwa-jp / heisei-jp）では元年 1 を「元」と表記する
     * @param {number} eraYear - 元号の年
     * @param {boolean} useGanText - 1 年を「元」と書くか
     * @returns {string} 年の文字列
     */
    function formatEraYearText(eraYear, useGanText) {
        if (useGanText && eraYear === 1) return "元";
        return String(eraYear);
    }

    /**
     * 元の数字テキストが 2 桁ゼロ詰め（"05" 等）なら、新しい数字も同じ桁数に揃える。
     * 新しい文字列が数字以外（例：「元」）の場合は揃えない
     * @param {string} oldText - 元の数字テキスト
     * @param {number|string} newValue - 新しい値
     * @returns {string} 桁をそろえた文字列
     */
    function applyZeroPaddingFrom(oldText, newValue) {
        var newStr = String(newValue);
        if (oldText.length === 2 && oldText.charAt(0) === '0' && newStr.length === 1 && /^[0-9]$/.test(newStr)) {
            return "0" + newStr;
        }
        return newStr;
    }

    /**
     * 実在する日付かを判定する
     * @param {number} year - 年
     * @param {number} month - 月
     * @param {number} day - 日
     * @returns {boolean} 実在すれば true
     */
    function isRealDate(year, month, day) {
        var checkDate = new Date(year, month - 1, day);
        return checkDate.getFullYear() === year &&
            checkDate.getMonth() === month - 1 &&
            checkDate.getDate() === day;
    }

    /**
     * 元号形式の日付が元号の有効範囲内かを判定する
     * @param {Object} parsedDate - detectFormatAndParse() の戻り値
     * @returns {boolean} 範囲内（または元号なし）なら true
     */
    function isEraDateValid(parsedDate) {
        if (!parsedDate) return false;
        var checkDate = new Date(parsedDate.year, parsedDate.month - 1, parsedDate.day);
        if (parsedDate.era === 'reiwa') return checkDate >= REIWA_START_DATE;
        if (parsedDate.era === 'heisei') return checkDate >= HEISEI_START_DATE && checkDate < HEISEI_END_DATE_EXCLUSIVE;
        return true;
    }

    /* 日付の形式ごとの解析規則（元号系を先に並べて優先させる）。
       hasPrefix: 令和・平成・R・H の接頭辞があるか / separator: 「.」「/」区切りの形式の区切り文字（年月日の漢字で区切る形式は null）
       Parse rules per date format (era formats first). The suffix group keeps any weekday text */
    var DATE_PARSE_RULES = [
        { format: 'reiwa-jp', era: 'reiwa', hasPrefix: true, separator: null,
            pattern: /^(令和)([0-9]{1,2})(年)(0?[1-9]|1[0-2])(月)(0?[1-9]|[12][0-9]|3[01])(日)(.*)$/ },
        { format: 'heisei-jp', era: 'heisei', hasPrefix: true, separator: null,
            pattern: /^(平成)([0-9]{1,2})(年)(0?[1-9]|1[0-2])(月)(0?[1-9]|[12][0-9]|3[01])(日)(.*)$/ },
        { format: 'r-dot', era: 'reiwa', hasPrefix: true, separator: ".",
            pattern: /^(R)([0-9]{1,2})\.(0?[1-9]|1[0-2])\.(0?[1-9]|[12][0-9]|3[01])(.*)$/ },
        { format: 'r-slash', era: 'reiwa', hasPrefix: true, separator: "/",
            pattern: /^(R)([0-9]{1,2})\/(0?[1-9]|1[0-2])\/(0?[1-9]|[12][0-9]|3[01])(.*)$/ },
        { format: 'h-dot', era: 'heisei', hasPrefix: true, separator: ".",
            pattern: /^(H)([0-9]{1,2})\.(0?[1-9]|1[0-2])\.(0?[1-9]|[12][0-9]|3[01])(.*)$/ },
        { format: 'h-slash', era: 'heisei', hasPrefix: true, separator: "/",
            pattern: /^(H)([0-9]{1,2})\/(0?[1-9]|1[0-2])\/(0?[1-9]|[12][0-9]|3[01])(.*)$/ },
        { format: 'jp', era: null, hasPrefix: false, separator: null,
            pattern: /^([0-9]{4})(年)(0?[1-9]|1[0-2])(月)(0?[1-9]|[12][0-9]|3[01])(日)(.*)$/ },
        { format: 'dot', era: null, hasPrefix: false, separator: ".",
            pattern: /^([0-9]{4})\.(0?[1-9]|1[0-2])\.(0?[1-9]|[12][0-9]|3[01])(.*)$/ },
        { format: 'slash', era: null, hasPrefix: false, separator: "/",
            pattern: /^([0-9]{4})\/(0?[1-9]|1[0-2])\/(0?[1-9]|[12][0-9]|3[01])(.*)$/ }
    ];

    /**
     * 年の文字列を西暦の数値にする（元号なら換算）
     * @param {string|null} era - 'reiwa' / 'heisei' / null
     * @param {string} yearText - 年の文字列
     * @returns {number} 西暦
     */
    function toGregorianYear(era, yearText) {
        var yearValue = parseInt(yearText, 10);
        if (era === 'reiwa') return reiwaToGregorian(yearValue);
        if (era === 'heisei') return heiseiToGregorian(yearValue);
        return yearValue;
    }

    /**
     * マッチ文字列を解析し、形式種別と各構成要素を返す。
     * 戻り値の主要フィールド：
     *   format: 'jp' | 'dot' | 'slash' | 'reiwa-jp' | 'r-dot' | 'r-slash' | 'heisei-jp' | 'h-dot' | 'h-slash'
     *   era:    'reiwa' | 'heisei' | null
     *   year, month, day: 西暦の数値（令和・平成形式の場合は変換後）
     *   parts:  元テキストの分解（prefix, year, sep1, month, sep2, day, sep3, suffix）
     * @param {string} matchText - マッチした文字列
     * @returns {Object|null} 解析結果（どの形式にも当たらなければ null）
     */
    function detectFormatAndParse(matchText) {
        var sourceText = String(matchText);
        for (var r = 0; r < DATE_PARSE_RULES.length; r++) {
            var parseRule = DATE_PARSE_RULES[r];
            var m = sourceText.match(parseRule.pattern);
            if (!m) continue;

            /* 接頭辞があると、以降のグループ番号が1つずれる / A prefix shifts the later groups by one */
            var groupOffset = parseRule.hasPrefix ? 1 : 0;
            var prefixText = parseRule.hasPrefix ? m[1] : "";
            var dateParts;
            if (parseRule.separator) {
                dateParts = {
                    prefix: prefixText, year: m[1 + groupOffset], sep1: parseRule.separator, month: m[2 + groupOffset],
                    sep2: parseRule.separator, day: m[3 + groupOffset], sep3: "", suffix: m[4 + groupOffset] || ""
                };
            } else {
                dateParts = {
                    prefix: prefixText, year: m[1 + groupOffset], sep1: m[2 + groupOffset], month: m[3 + groupOffset],
                    sep2: m[4 + groupOffset], day: m[5 + groupOffset], sep3: m[6 + groupOffset], suffix: m[7 + groupOffset] || ""
                };
            }
            return {
                format: parseRule.format, era: parseRule.era,
                year: toGregorianYear(parseRule.era, dateParts.year),
                month: parseInt(dateParts.month, 10),
                day: parseInt(dateParts.day, 10),
                parts: dateParts
            };
        }
        return null;
    }

    // =========================================
    // 確認パネルの表示 / Check panel text
    // =========================================

    /**
     * 2 つの日付の差を日数で返す
     * @param {Date} fromDate - 基準の日付
     * @param {Date} toDate - 比べる日付
     * @returns {number} 日数の差
     */
    function getDaysDifference(fromDate, toDate) {
        var msPerDay = 1000 * 60 * 60 * 24;
        return Math.round((toDate.getTime() - fromDate.getTime()) / msPerDay);
    }

    /**
     * 日数差ラベル（符号付き、例：「+15日」）
     * @param {number} days - 日数の差
     * @returns {string} 表示用の文字列
     */
    function formatDaysDifference(days) {
        var sign = days > 0 ? "+" : "";
        return sign + days + "日";
    }

    /**
     * 曜日表示（例：金曜日）
     * @param {number} year - 年
     * @param {number} month - 月
     * @param {number} day - 日
     * @returns {string} 曜日（日付でなければ空文字）
     */
    function getWeekdayLabel(year, month, day) {
        if (isNaN(year) || isNaN(month) || isNaN(day)) return "";
        var checkDate = new Date(year, month - 1, day);
        if (isNaN(checkDate.getTime())) return "";
        return WEEKDAY_KANJI[checkDate.getDay()] + "曜日";
    }

    /**
     * 西暦+月日から和暦表記を返す（例：「令和8年」「平成元年」）。明治より前は空文字
     * @param {number} year - 年
     * @param {number} month - 月
     * @param {number} day - 日
     * @returns {string} 和暦表記
     */
    function formatEraLabel(year, month, day) {
        if (isNaN(year) || isNaN(month) || isNaN(day)) return "";
        var checkDate = new Date(year, month - 1, day);
        if (isNaN(checkDate.getTime())) return "";

        for (var i = 0; i < ERA_TABLE.length; i++) {
            if (checkDate >= ERA_TABLE[i].startDate) {
                return ERA_TABLE[i].name + formatEraYearText(year - ERA_TABLE[i].baseYear, true) + "年";
            }
        }
        return "";
    }

    // =========================================
    // 日付検索 / Searching dates
    // =========================================

    /*
       対象例：
         2026年5月8日 / 2026年5月8日金曜 / 2026年5月8日金曜日 / 2026年5月8日(金) / 2026年5月8日（金）
         2026.5.8 / 2026.5.8（金） / 2026/5/8 / 2026/5/8（金）
         令和8年5月8日 / 令和8年5月8日（金） / R8.5.8 / R8.5.8（金） / R8/5/8 / R8/5/8（金）
         英語曜日サフィックス（Fri / Friday など）も検出する。
       元号系を先に列挙して優先マッチさせる。
    */

    var MONTH_PATTERN = "(?:0?[1-9]|1[0-2])";
    var DAY_PATTERN = "(?:0?[1-9]|[12][0-9]|3[01])";

    var WEEKDAY_SUFFIX_PATTERN = "(?:[日月火水木金土]曜日|[日月火水木金土]曜|\\([日月火水木金土]\\)|（[日月火水木金土]）|Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sun|Mon|Tue|Wed|Thu|Fri|Sat)?";
    var DATE_SEARCH_REGEX = new RegExp(
        "令和[0-9]{1,2}年" + MONTH_PATTERN + "月" + DAY_PATTERN + "日" + WEEKDAY_SUFFIX_PATTERN +
        "|平成[0-9]{1,2}年" + MONTH_PATTERN + "月" + DAY_PATTERN + "日" + WEEKDAY_SUFFIX_PATTERN +
        "|R[0-9]{1,2}\\." + MONTH_PATTERN + "\\." + DAY_PATTERN + WEEKDAY_SUFFIX_PATTERN +
        "|R[0-9]{1,2}/" + MONTH_PATTERN + "/" + DAY_PATTERN + WEEKDAY_SUFFIX_PATTERN +
        "|H[0-9]{1,2}\\." + MONTH_PATTERN + "\\." + DAY_PATTERN + WEEKDAY_SUFFIX_PATTERN +
        "|H[0-9]{1,2}/" + MONTH_PATTERN + "/" + DAY_PATTERN + WEEKDAY_SUFFIX_PATTERN +
        "|[0-9]{4}年" + MONTH_PATTERN + "月" + DAY_PATTERN + "日" + WEEKDAY_SUFFIX_PATTERN +
        "|[0-9]{4}\\." + MONTH_PATTERN + "\\." + DAY_PATTERN + WEEKDAY_SUFFIX_PATTERN +
        "|[0-9]{4}/" + MONTH_PATTERN + "/" + DAY_PATTERN + WEEKDAY_SUFFIX_PATTERN,
        "g"
    );

    /**
     * テキストフレームを重複なく追加する
     * @param {TextFrame} textFrame - 追加するテキストフレーム
     * @param {TextFrame[]} accumulator - 追加先
     * @returns {void}
     */
    function addTextFrameOnce(textFrame, accumulator) {
        if (!textFrame || textFrame.typename !== "TextFrame") return;
        for (var frameIndex = 0; frameIndex < accumulator.length; frameIndex++) {
            if (accumulator[frameIndex] === textFrame) return;
        }
        accumulator.push(textFrame);
    }

    /**
     * 部分テキスト選択から親テキストフレームを取得する
     * @param {Object} selectedText - TextRange などの文字の選択
     * @returns {TextFrame|null} 親のテキストフレーム
     */
    function getParentTextFrameFromTextSelection(selectedText) {
        var currentItem = selectedText;
        /* parent が自身を指すような壊れた参照に備え、辿る階層数に上限を設ける */
        var MAX_PARENT_DEPTH = 10;
        for (var depth = 0; depth < MAX_PARENT_DEPTH && currentItem; depth++) {
            try {
                if (currentItem.typename === "TextFrame") return currentItem;
                currentItem = currentItem.parent;
            } catch (eParentTextFrame) {
                break;
            }
        }
        return null;
    }

    /**
     * 選択オブジェクト（およびその子）からテキストフレームを再帰的に収集する
     * @param {Object[]} selectedItems - 選択中のオブジェクト
     * @param {TextFrame[]} accumulator - 追加先
     * @returns {void}
     */
    function collectTextFramesFromItems(selectedItems, accumulator) {
        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            if (!selectedItem) continue;
            if (selectedItem.typename === "TextFrame") {
                addTextFrameOnce(selectedItem, accumulator);
            } else if (selectedItem.typename === "GroupItem") {
                collectTextFramesFromItems(selectedItem.pageItems, accumulator);
            } else if (selectedItem.typename === "TextRange" || selectedItem.typename === "InsertionPoint" || selectedItem.typename === "Character") {
                addTextFrameOnce(getParentTextFrameFromTextSelection(selectedItem), accumulator);
            }
        }
    }

    /**
     * 検索対象のテキストフレームを集める（選択があれば選択範囲内、なければドキュメント全体）
     * @param {Document} doc - 対象のドキュメント
     * @returns {TextFrame[]} 検索対象のテキストフレーム
     */
    function collectTargetTextFrames(doc) {
        var textFrames = [];
        var docSelection = doc.selection;
        if (docSelection && docSelection.length > 0) {
            /* 選択あり：選択範囲内のテキストフレームのみ */
            collectTextFramesFromItems(docSelection, textFrames);
        } else {
            /* 選択なし：ドキュメント全体 */
            for (var i = 0; i < doc.textFrames.length; i++) {
                textFrames.push(doc.textFrames[i]);
            }
        }
        return textFrames;
    }

    /**
     * テキストフレームから実在する日付を探す
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame[]} textFrames - 検索対象のテキストフレーム
     * @returns {Object[]} 見つかった日付（frameIndex / text / matchIndex / artboardIndex / parsed など）
     */
    function findDateMatches(doc, textFrames) {
        var foundMatches = [];
        for (var frameIdx = 0; frameIdx < textFrames.length; frameIdx++) {
            var targetFrame = textFrames[frameIdx];
            var frameContent = "";

            try {
                frameContent = targetFrame.contents;
            } catch (eRead) {
                continue;
            }

            var artboardIdx = -1;
            try {
                artboardIdx = getArtboardIndex(doc, targetFrame);
            } catch (eArtboard) {
                artboardIdx = -1;
            }

            var match;
            DATE_SEARCH_REGEX.lastIndex = 0;

            while ((match = DATE_SEARCH_REGEX.exec(frameContent)) !== null) {
                /* 直前が数字、かつマッチが数字始まり（YYYY 系）の場合は、5 桁以上の連続数字を切り出した
                   誤検出とみなしてスキップ。R/H/令和/平成 始まりはこの判定の対象外 */
                if (match.index > 0 &&
                    /[0-9]/.test(frameContent.charAt(match.index - 1)) &&
                    /^[0-9]/.test(match[0])) {
                    continue;
                }
                var parsedDate = detectFormatAndParse(match[0]);
                if (!parsedDate) continue;
                if (!isRealDate(parsedDate.year, parsedDate.month, parsedDate.day)) continue;
                if (!isEraDateValid(parsedDate)) continue;
                foundMatches.push({
                    frameIndex: frameIdx,
                    text: match[0],
                    matchIndex: match.index,
                    artboardIndex: artboardIdx,
                    parsed: parsedDate,
                    weekdaySuffixStyle: getWeekdaySuffixStyle(parsedDate),
                    weekdayPairs: null,
                    checkbox: null
                });
            }
        }
        return foundMatches;
    }

    /**
     * 日付マッチごとに、その TextFrame の直接の親 GroupItem を見て、
     * 同じ親グループ内の「曜日のみ」TextFrame を曜日ペアとして関連付ける（match.weekdayPairs に入れる）。
     * 同一グループに複数の日付マッチがある場合は対応が曖昧なのでペアリングしない。
     * @param {Object[]} foundMatches - findDateMatches() の戻り値
     * @param {TextFrame[]} textFrames - 検索対象のテキストフレーム
     * @returns {void}
     */
    function pairWeekdayFrames(foundMatches, textFrames) {
        for (var pairingMatchIndex = 0; pairingMatchIndex < foundMatches.length; pairingMatchIndex++) {
            var pairingMatch = foundMatches[pairingMatchIndex];
            var dateFrame = textFrames[pairingMatch.frameIndex];
            var parentGroup = null;
            try {
                if (dateFrame.parent && dateFrame.parent.typename === 'GroupItem') {
                    parentGroup = dateFrame.parent;
                }
            } catch (eParentGroup) { parentGroup = null; }
            if (!parentGroup) continue;

            var hasOtherDateInSameGroup = false;
            for (var otherMatchIndex = 0; otherMatchIndex < foundMatches.length; otherMatchIndex++) {
                if (otherMatchIndex === pairingMatchIndex) continue;
                try {
                    if (textFrames[foundMatches[otherMatchIndex].frameIndex].parent === parentGroup) {
                        hasOtherDateInSameGroup = true;
                        break;
                    }
                } catch (eOtherMatch) { }
            }
            if (hasOtherDateInSameGroup) continue;

            var weekdayPairs = [];
            for (var childItemIndex = 0; childItemIndex < parentGroup.pageItems.length; childItemIndex++) {
                var childItem;
                try { childItem = parentGroup.pageItems[childItemIndex]; } catch (eChildItem) { continue; }
                if (childItem === dateFrame) continue;
                if (childItem.typename !== 'TextFrame') continue;
                try {
                    var weekdayStyle = detectWeekdayOnlyFrame(childItem.contents);
                    if (weekdayStyle) {
                        weekdayPairs.push({ frame: childItem, style: weekdayStyle });
                    }
                } catch (eWeekdayContent) { }
            }
            if (weekdayPairs.length > 0) {
                pairingMatch.weekdayPairs = weekdayPairs;
            }
        }
    }

    /**
     * 曜日ドロップダウンの初期選択を決める（最初の日付の曜日サフィックス、なければ曜日ペアの形式）
     * 日付サフィックスは単独漢字曜日を許可しないため 'kanji' は曜日のみフレームからだけ拾う
     * @param {Object[]} foundMatches - 見つかった日付
     * @returns {string} WEEKDAY_VALUES の値
     */
    function detectInitialWeekdayChoice(foundMatches) {
        for (var matchIndex = 0; matchIndex < foundMatches.length; matchIndex++) {
            var matchInfo = foundMatches[matchIndex];
            if (matchInfo.weekdaySuffixStyle) return matchInfo.weekdaySuffixStyle;
            if (matchInfo.weekdayPairs && matchInfo.weekdayPairs.length > 0) return matchInfo.weekdayPairs[0].style;
        }
        return 'none';
    }

    // =========================================
    // 置換 / Replacement
    // =========================================

    /**
     * テキストフレーム内の一致範囲を書式を保持したまま置換する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} matchStart - 置換を始める文字位置
     * @param {number} oldLen - 置換する文字数
     * @param {string} newText - 新しい文字列
     * @returns {number} 実行した文字操作数（プレビュー巻き戻し用）
     */
    function replaceMatchPreserveStyle(textFrame, matchStart, oldLen, newText) {
        if (oldLen <= 0) return 0;
        var newLen = newText.length;
        try {
            if (textFrame.contents.substr(matchStart, oldLen) === newText) return 0;
        } catch (eSameTextCheck) { }
        var minLen = Math.min(oldLen, newLen);
        /* 伸長時は最後の元文字を「最後の新文字＋追加分」で一度だけ書き換えるため、ループは手前で止める */
        var loopEnd = (newLen > oldLen) ? minLen - 1 : minLen;
        var opCount = 0;

        /* 重複部分は文字単位で書き換え。各 character.contents の代入は元の文字属性を維持 */
        for (var i = 0; i < loopEnd; i++) {
            textFrame.characters[matchStart + i].contents = newText.charAt(i);
            opCount++;
        }

        if (newLen > oldLen) {
            /* 末尾文字に追加分の文字列を含めると、元の書式を引き継いだまま文字を増やせる */
            var lastIdx = matchStart + oldLen - 1;
            textFrame.characters[lastIdx].contents = newText.charAt(oldLen - 1) + newText.substring(oldLen);
            opCount++;
        } else if (newLen < oldLen) {
            /* 余った古い文字を末尾から削除 */
            for (var ri = oldLen - 1; ri >= newLen; ri--) {
                textFrame.characters[matchStart + ri].remove();
                opCount++;
            }
        }
        return opCount;
    }

    /**
     * 指定した形式の曜日文字列を作る
     * @param {string} weekdayChoice - WEEKDAY_VALUES の値
     * @param {number} year - 年
     * @param {number} month - 月
     * @param {number} day - 日
     * @returns {string} 曜日の文字列（'none' や不正な日付なら空文字）
     */
    function buildExplicitWeekday(weekdayChoice, year, month, day) {
        if (weekdayChoice === 'none') return "";
        var checkDate = new Date(year, month - 1, day);
        if (isNaN(checkDate.getTime())) return "";
        var weekdayIndex = checkDate.getDay();
        switch (weekdayChoice) {
            case 'kanji': return WEEKDAY_KANJI[weekdayIndex];
            case 'medium': return WEEKDAY_KANJI[weekdayIndex] + "曜";
            case 'long': return WEEKDAY_KANJI[weekdayIndex] + "曜日";
            case 'full-paren': return "（" + WEEKDAY_KANJI[weekdayIndex] + "）";
            case 'half-paren': return "(" + WEEKDAY_KANJI[weekdayIndex] + ")";
            case 'en-short': return WEEKDAY_EN_SHORT[weekdayIndex];
            case 'en-full': return WEEKDAY_EN_FULL[weekdayIndex];
        }
        return "";
    }

    /**
     * @typedef {object} ReplaceOptions
     * @property {number} newYear - 置換後の年（西暦）
     * @property {number} newMonth - 置換後の月
     * @property {number} newDay - 置換後の日
     * @property {string} formatChoice - FORMAT_VALUES の値
     * @property {string} weekdayChoice - WEEKDAY_VALUES の値
     * @property {boolean} preserveNumberFormat - 数字の書式を保持するか
     * @property {string} explicitWeekdayText - 付ける曜日の文字列
     */

    /**
     * 元の形式での新しい年の文字列を作る（元号なら換算し、ゼロ詰めも元に合わせる）
     * @param {Object} parsedDate - 元の日付の解析結果
     * @param {number} newYear - 置換後の年（西暦）
     * @returns {string} 年の文字列
     */
    function buildNewYearText(parsedDate, newYear) {
        var newYearText;
        if (parsedDate.era === 'reiwa') newYearText = formatEraYearText(gregorianToReiwa(newYear), parsedDate.format === 'reiwa-jp');
        else if (parsedDate.era === 'heisei') newYearText = formatEraYearText(gregorianToHeisei(newYear), parsedDate.format === 'heisei-jp');
        else newYearText = String(newYear);
        return applyZeroPaddingFrom(parsedDate.parts.year, newYearText);
    }

    /**
     * 元の形式（区切り文字・元号）を保った置換後の文字列を作る
     * @param {Object} matchInfo - 見つかった日付
     * @param {ReplaceOptions} replaceOptions - 置換の設定
     * @returns {string|null} 置換後の文字列
     */
    function buildPreservedFormatText(matchInfo, replaceOptions) {
        var parsedDate = matchInfo.parsed;
        if (!parsedDate) return null;
        var oldParts = parsedDate.parts;
        var replacedText =
            oldParts.prefix +
            buildNewYearText(parsedDate, replaceOptions.newYear) +
            oldParts.sep1 +
            applyZeroPaddingFrom(oldParts.month, replaceOptions.newMonth) +
            oldParts.sep2 +
            applyZeroPaddingFrom(oldParts.day, replaceOptions.newDay) +
            oldParts.sep3;
        /* 曜日ドロップダウンの選択を常に反映（「なし」なら元の suffix を削除） */
        replacedText += replaceOptions.explicitWeekdayText;
        return replacedText;
    }

    /**
     * 選んだ形式で置換後の文字列を作る
     * @param {string} format - FORMAT_VALUES の値（preserve 以外）
     * @param {ReplaceOptions} replaceOptions - 置換の設定
     * @returns {string|null} 置換後の文字列
     */
    function buildExplicitFormatText(format, replaceOptions) {
        var newYear = replaceOptions.newYear;
        var newMonth = replaceOptions.newMonth;
        var newDay = replaceOptions.newDay;
        var weekdayText = replaceOptions.explicitWeekdayText;
        var reiwaYear = gregorianToReiwa(newYear);
        var heiseiYear = gregorianToHeisei(newYear);
        var reiwaJpYearText = formatEraYearText(reiwaYear, true);
        var heiseiJpYearText = formatEraYearText(heiseiYear, true);
        switch (format) {
            case 'jp': return newYear + "年" + newMonth + "月" + newDay + "日" + weekdayText;
            case 'jp-md': return newMonth + "月" + newDay + "日" + weekdayText;
            case 'dot': return newYear + "." + newMonth + "." + newDay + weekdayText;
            case 'dot-md': return newMonth + "." + newDay + weekdayText;
            case 'slash': return newYear + "/" + newMonth + "/" + newDay + weekdayText;
            case 'slash-md': return newMonth + "/" + newDay + weekdayText;
            case 'reiwa-jp': return "令和" + reiwaJpYearText + "年" + newMonth + "月" + newDay + "日" + weekdayText;
            case 'r-dot': return "R" + reiwaYear + "." + newMonth + "." + newDay + weekdayText;
            case 'r-slash': return "R" + reiwaYear + "/" + newMonth + "/" + newDay + weekdayText;
            case 'heisei-jp': return "平成" + heiseiJpYearText + "年" + newMonth + "月" + newDay + "日" + weekdayText;
            case 'h-dot': return "H" + heiseiYear + "." + newMonth + "." + newDay + weekdayText;
            case 'h-slash': return "H" + heiseiYear + "/" + newMonth + "/" + newDay + weekdayText;
        }
        return null;
    }

    /**
     * 数字の書式を保つため、年・月・日などの部分ごとの置換区間を作る
     * @param {Object} matchInfo - 見つかった日付
     * @param {ReplaceOptions} replaceOptions - 置換の設定
     * @returns {{offset: number, oldLen: number, newText: string}[]|null} 置換区間（マッチ先頭からの位置）
     */
    function buildReplacementSegments(matchInfo, replaceOptions) {
        var parsedDate = matchInfo.parsed;
        if (!parsedDate) return null;
        var oldParts = parsedDate.parts;
        var segments = [];
        var segmentPos = 0;
        function pushSegment(oldText, newText) {
            if (oldText.length === 0) return;
            segments.push({ offset: segmentPos, oldLen: oldText.length, newText: newText });
            segmentPos += oldText.length;
        }
        var hasJpSuffixSlot = (parsedDate.format === 'jp' || parsedDate.format === 'reiwa-jp' || parsedDate.format === 'heisei-jp');

        pushSegment(oldParts.prefix, oldParts.prefix);
        pushSegment(oldParts.year, buildNewYearText(parsedDate, replaceOptions.newYear));
        pushSegment(oldParts.sep1, oldParts.sep1);
        pushSegment(oldParts.month, applyZeroPaddingFrom(oldParts.month, replaceOptions.newMonth));
        pushSegment(oldParts.sep2, oldParts.sep2);

        if (hasJpSuffixSlot) {
            pushSegment(oldParts.day, applyZeroPaddingFrom(oldParts.day, replaceOptions.newDay));
            if (oldParts.suffix.length > 0) {
                /* 元 suffix を曜日ドロップダウンの選択で置換（「なし」のときは空文字で削除） */
                pushSegment(oldParts.sep3, oldParts.sep3);
                pushSegment(oldParts.suffix, replaceOptions.explicitWeekdayText);
            } else {
                /* 元 suffix なし：sep3（"日"）末尾に曜日を伸ばす */
                pushSegment(oldParts.sep3, oldParts.sep3 + replaceOptions.explicitWeekdayText);
            }
        } else {
            /* dot / slash / r-dot / r-slash：曜日サフィックスを含めて日付全体を1セグメントとして置換する */
            var preservedText = buildPreservedFormatText(matchInfo, replaceOptions);
            if (preservedText !== null) {
                segments = [{ offset: 0, oldLen: matchInfo.text.length, newText: preservedText }];
            }
        }
        return segments;
    }

    /**
     * 1つのテキストフレーム内のマッチを後ろから置換する（途中で例外が出ても、それまでの操作数は replaceStats に残る）
     * @param {TextFrame} targetFrame - 対象のテキストフレーム
     * @param {Object[]} frameMatches - このフレームのマッチ（matchIndex の降順）
     * @param {ReplaceOptions} replaceOptions - 置換の設定
     * @param {{operations: number, errors: number}} replaceStats - 文字操作数を足していく集計
     * @returns {void}
     */
    function replaceFrameMatches(targetFrame, frameMatches, replaceOptions, replaceStats) {
        for (var orderIdx = 0; orderIdx < frameMatches.length; orderIdx++) {
            var currentMatch = frameMatches[orderIdx];
            if (replaceOptions.formatChoice === 'preserve' && replaceOptions.preserveNumberFormat) {
                var segments = buildReplacementSegments(currentMatch, replaceOptions);
                if (!segments) continue;
                for (var segIdx = segments.length - 1; segIdx >= 0; segIdx--) {
                    var segment = segments[segIdx];
                    replaceStats.operations += replaceMatchPreserveStyle(targetFrame, currentMatch.matchIndex + segment.offset, segment.oldLen, segment.newText);
                }
            } else {
                var fullText = (replaceOptions.formatChoice === 'preserve')
                    ? buildPreservedFormatText(currentMatch, replaceOptions)
                    : buildExplicitFormatText(replaceOptions.formatChoice, replaceOptions);
                if (fullText === null) continue;
                replaceStats.operations += replaceMatchPreserveStyle(targetFrame, currentMatch.matchIndex, currentMatch.text.length, fullText);
            }
        }
    }

    /**
     * チェックされたマッチに対して置換を実行する
     * @param {Object[]} foundMatches - 見つかった日付
     * @param {TextFrame[]} textFrames - 検索対象のテキストフレーム
     * @param {ReplaceOptions} replaceOptions - 置換の設定
     * @returns {{operations: number, errors: number, selectedCount: number}} 文字操作数・エラー件数・対象件数
     */
    function performReplacement(foundMatches, textFrames, replaceOptions) {
        /* チェック済みマッチをフレーム単位でまとめる */
        var matchesByFrame = {};
        var selectedCount = 0;
        for (var resultIdx = 0; resultIdx < foundMatches.length; resultIdx++) {
            if (!foundMatches[resultIdx].checkbox || !foundMatches[resultIdx].checkbox.value) continue;
            var targetFrameIdx = foundMatches[resultIdx].frameIndex;
            if (!matchesByFrame[targetFrameIdx]) matchesByFrame[targetFrameIdx] = [];
            matchesByFrame[targetFrameIdx].push(foundMatches[resultIdx]);
            selectedCount++;
        }

        var replaceStats = { operations: 0, errors: 0 };

        for (var frameIdxKey in matchesByFrame) {
            if (!matchesByFrame.hasOwnProperty(frameIdxKey)) continue;
            var frameMatches = matchesByFrame[frameIdxKey];
            /* 後ろから置換すれば前方の matchIndex はずれない */
            frameMatches.sort(function (firstMatch, secondMatch) { return secondMatch.matchIndex - firstMatch.matchIndex; });

            try {
                replaceFrameMatches(textFrames[parseInt(frameIdxKey, 10)], frameMatches, replaceOptions, replaceStats);
            } catch (eReplace) {
                replaceStats.errors++;
            }
        }

        /*
           チェック済みマッチに紐づく曜日ペア（同一グループ内の曜日のみフレーム）を更新。
           曜日ドロップダウンが「なし」の場合は、空文字化せず連動フレームを更新しない。
           各ペアフレームは独立した TextFrame のままで、日付フレームには連結しない。
        */
        if (replaceOptions.weekdayChoice !== 'none') {
            for (var pairMatchIndex = 0; pairMatchIndex < foundMatches.length; pairMatchIndex++) {
                var dateMatchWithPairs = foundMatches[pairMatchIndex];
                if (!dateMatchWithPairs.checkbox || !dateMatchWithPairs.checkbox.value) continue;
                if (!dateMatchWithPairs.weekdayPairs) continue;
                for (var weekdayPairIndex = 0; weekdayPairIndex < dateMatchWithPairs.weekdayPairs.length; weekdayPairIndex++) {
                    var weekdayPair = dateMatchWithPairs.weekdayPairs[weekdayPairIndex];
                    try {
                        /* 元のスタイルを維持しつつ、新しい曜日に置換 */
                        var newWeekdayText = buildExplicitWeekday(weekdayPair.style, replaceOptions.newYear, replaceOptions.newMonth, replaceOptions.newDay);
                        if (newWeekdayText === "") continue;
                        var oldWeekdayText = String(weekdayPair.frame.contents);
                        if (oldWeekdayText === newWeekdayText) continue;
                        replaceStats.operations += replaceMatchPreserveStyle(weekdayPair.frame, 0, oldWeekdayText.length, newWeekdayText);
                    } catch (eWeekdayPair) {
                        replaceStats.errors++;
                    }
                }
            }
        }

        return { operations: replaceStats.operations, errors: replaceStats.errors, selectedCount: selectedCount };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 見つかった日付をアートボードごとにまとめる（アートボード外の -1 は末尾）
     * @param {Object[]} foundMatches - 見つかった日付
     * @returns {{artboardIndexes: number[], matchesByArtboard: Object}} 表示順のアートボード番号と、番号ごとのマッチ
     */
    function groupMatchesByArtboard(foundMatches) {
        var matchesByArtboard = {};
        var artboardIndexes = [];
        for (var matchIdx = 0; matchIdx < foundMatches.length; matchIdx++) {
            var artboardKey = String(foundMatches[matchIdx].artboardIndex);
            if (!matchesByArtboard[artboardKey]) {
                matchesByArtboard[artboardKey] = [];
                artboardIndexes.push(foundMatches[matchIdx].artboardIndex);
            }
            matchesByArtboard[artboardKey].push(foundMatches[matchIdx]);
        }
        /* -1（アートボード外）は末尾に */
        artboardIndexes.sort(function (firstIndex, secondIndex) {
            if (firstIndex === -1 && secondIndex !== -1) return 1;
            if (secondIndex === -1 && firstIndex !== -1) return -1;
            return firstIndex - secondIndex;
        });
        return { artboardIndexes: artboardIndexes, matchesByArtboard: matchesByArtboard };
    }

    /**
     * ［見つかった日付］パネルを追加し、日付ごとのチェックボックスを作る（match.checkbox に入れる）
     * @param {Window} dateDialog - 追加先のダイアログ
     * @param {Document} doc - 対象のドキュメント
     * @param {Object[]} foundMatches - 見つかった日付
     * @returns {Checkbox[]} 作ったチェックボックス
     */
    function addFoundDatesPanel(dateDialog, doc, foundMatches) {
        var dateCheckboxes = [];
        var artboardGroups = groupMatchesByArtboard(foundMatches);

        var foundDatesPanel = dateDialog.add("panel", undefined, "見つかった日付（" + foundMatches.length + "）");
        setupPanel(foundDatesPanel, PANEL_SPACING);
        foundDatesPanel.minimumSize.width = FOUND_PANEL_MIN_WIDTH;

        for (var keyIdx = 0; keyIdx < artboardGroups.artboardIndexes.length; keyIdx++) {
            var currentArtboardIdx = artboardGroups.artboardIndexes[keyIdx];
            var artboardHeaderText;
            if (currentArtboardIdx === -1) {
                artboardHeaderText = "[アートボード外]";
            } else {
                artboardHeaderText = "[アートボード " + (currentArtboardIdx + 1) + "：" +
                    doc.artboards[currentArtboardIdx].name + "]";
            }
            /* アートボード見出しは下に少し余白を取る */
            var artboardHeaderGroup = foundDatesPanel.add("group");
            artboardHeaderGroup.orientation = "row";
            artboardHeaderGroup.alignment = "left";
            artboardHeaderGroup.margins = [0, 4, 0, 6];
            artboardHeaderGroup.add("statictext", undefined, artboardHeaderText);

            var artboardMatches = artboardGroups.matchesByArtboard[String(currentArtboardIdx)];
            for (var itemIdx = 0; itemIdx < artboardMatches.length; itemIdx++) {
                /* 各日付行は左マージンでインデント */
                var checkboxRow = foundDatesPanel.add("group");
                checkboxRow.orientation = "row";
                checkboxRow.alignment = "left";
                checkboxRow.margins = [DATE_ROW_INDENT, 0, 0, 0];
                var checkboxText = artboardMatches[itemIdx].text;
                var pairCount = (artboardMatches[itemIdx].weekdayPairs ? artboardMatches[itemIdx].weekdayPairs.length : 0);
                if (pairCount > 0) {
                    checkboxText += "  ＋曜日連動";
                }
                var dateCheckbox = checkboxRow.add("checkbox", undefined, checkboxText);
                dateCheckbox.value = true;
                /* チェックボックスのラベル切れ対策 */
                dateCheckbox.preferredSize.width = DATE_CHECKBOX_WIDTH;
                dateCheckbox.helpTip = (pairCount > 0)
                    ? ("Option＋クリックで全項目を一括切替\n同一グループ内の曜日フレーム " + pairCount + " 件も連動して更新されます\n曜日が「なし」の場合は連動フレームを更新しません")
                    : "Option＋クリックで全項目を一括切替";
                artboardMatches[itemIdx].checkbox = dateCheckbox;
                dateCheckboxes.push(dateCheckbox);
            }
        }
        return dateCheckboxes;
    }

    /**
     * 年・月・日の入力欄を1つ追加する
     * @param {Group} dateInputGroup - 追加先
     * @param {number} initialValue - 初期値
     * @param {number} fieldChars - 入力欄の幅（文字数）
     * @param {string} unitText - 右に添える「年」「月」「日」
     * @returns {EditText} 追加した入力欄
     */
    function addDateField(dateInputGroup, initialValue, fieldChars, unitText) {
        var dateField = dateInputGroup.add("edittext", undefined, String(initialValue));
        dateField.characters = fieldChars;
        dateField.helpTip = ARROW_KEY_HELP;
        dateInputGroup.add("statictext", undefined, unitText);
        return dateField;
    }

    /**
     * 確認パネルに「項目名＋値」の行を追加する
     * @param {Panel} checkPanel - 追加先
     * @param {string} rowLabelText - 項目名
     * @param {number} valueWidth - 値の欄の幅
     * @returns {StaticText} 値の欄
     */
    function addCheckRow(checkPanel, rowLabelText, valueWidth) {
        var checkRow = checkPanel.add("group");
        checkRow.orientation = "row";
        var rowLabel = checkRow.add("statictext", undefined, rowLabelText);
        rowLabel.preferredSize.width = CHECK_LABEL_WIDTH;
        var valueText = checkRow.add("statictext", undefined, "");
        valueText.preferredSize.width = valueWidth;
        return valueText;
    }

    /**
     * ダイアログを組み立てる（イベントはまだ付けない）
     * @param {Document} doc - 対象のドキュメント
     * @param {Object[]} foundMatches - 見つかった日付
     * @returns {Object} ダイアログと各コントロール
     */
    function buildDialog(doc, foundMatches) {
        var dateDialog = new Window("dialog", "日付を検索・置換 " + SCRIPT_VERSION);
        dateDialog.orientation = "column";
        dateDialog.alignChildren = "fill";

        var dateCheckboxes = [];
        if (foundMatches.length === 0) {
            dateDialog.add("statictext", undefined, "日付は見つかりませんでした。");
        } else {
            dateCheckboxes = addFoundDatesPanel(dateDialog, doc, foundMatches);
        }

        /* 置換後の日付入力パネル：［2026］年［5］月［8］日 形式。初期値は今日 */
        var today = new Date();
        var replacementPanel = dateDialog.add("panel", undefined, "置換後の日付");
        setupPanel(replacementPanel, PANEL_SPACING);

        var dateInputGroup = replacementPanel.add("group");
        dateInputGroup.orientation = "row";
        var yearInput = addDateField(dateInputGroup, today.getFullYear(), 5, "年");
        var monthInput = addDateField(dateInputGroup, today.getMonth() + 1, 3, "月");
        var dayInput = addDateField(dateInputGroup, today.getDate(), 3, "日");

        var formatRow = replacementPanel.add("group");
        formatRow.orientation = "row";
        formatRow.add("statictext", undefined, "フォーマット：");
        var formatDropdown = formatRow.add("dropdownlist", undefined, FORMAT_LABELS);
        formatDropdown.selection = 0;
        formatDropdown.helpTip = "「元の形式を保持」：各マッチの区切り文字・元号を維持\n曜日表記は右の曜日ドロップダウンの選択を反映\nそれ以外：すべてのマッチを選択した形式に統一";

        var weekdayDropdown = formatRow.add("dropdownlist", undefined, WEEKDAY_LABELS);
        weekdayDropdown.helpTip = "出力に付与する曜日表記。「元の形式を保持」選択時もこの選択を反映し、「なし」で日付内の曜日表記を削除。連動曜日フレームは更新しません";

        /* 最初に見つかった置換対象を基準に、曜日サフィックスまたは曜日のみフレームの形式を初期選択にする */
        var initialWeekdayChoice = detectInitialWeekdayChoice(foundMatches);
        var initialWeekdayIndex = 0;
        for (var wi = 0; wi < WEEKDAY_VALUES.length; wi++) {
            if (WEEKDAY_VALUES[wi] === initialWeekdayChoice) { initialWeekdayIndex = wi; break; }
        }
        weekdayDropdown.selection = initialWeekdayIndex;

        /* 数字（年・月・日）の文字書式を保持するか。OFF にするとマッチ範囲を一括置換し、
           新しい文字はマッチ先頭の書式に統一される */
        var preserveNumberFormatCheckbox = replacementPanel.add("checkbox", undefined, "数字の書式を保持する");
        preserveNumberFormatCheckbox.value = true;
        preserveNumberFormatCheckbox.helpTip = "ON：年・月・日それぞれの元の文字書式（フォント・サイズ・色など）を維持\nOFF：マッチ範囲全体を一括置換し、書式は先頭文字に揃える\n※「元の形式を保持」選択時のみ有効";

        /* プレビュー：ON で現在の入力をドキュメントに反映し、ダイアログを開いたまま結果を確認できる。
           入力変更時には自動で更新（巻き戻し→再適用）。OFF にすると元に戻す */
        var previewCheckbox = replacementPanel.add("checkbox", undefined, "プレビュー");
        previewCheckbox.value = false;
        previewCheckbox.helpTip = "ON でドキュメントに即時反映。入力やチェックを変更すると自動更新。\nOFF・キャンセルで元に戻す";

        /* 確認パネル：和暦・曜日・日数差のプレビュー（置換対象外） */
        var checkPanel = dateDialog.add("panel", undefined, "置換内容の確認");
        setupPanel(checkPanel, PANEL_SPACING);
        var eraLabel = addCheckRow(checkPanel, "和暦：", 120);
        var weekdayLabel = addCheckRow(checkPanel, "置換後の曜日：", 80);
        var daysDiffLabel = addCheckRow(checkPanel, "最初の日付との差：", 120);

        /* OK / キャンセル ボタン */
        var btnRowGroup = dateDialog.add("group");
        btnRowGroup.alignment = "right";
        var btnCancel = btnRowGroup.add("button", undefined, "キャンセル");
        var btnOK = btnRowGroup.add("button", undefined, "OK");

        return {
            dialog: dateDialog,
            dateCheckboxes: dateCheckboxes,
            yearInput: yearInput,
            monthInput: monthInput,
            dayInput: dayInput,
            formatDropdown: formatDropdown,
            weekdayDropdown: weekdayDropdown,
            preserveNumberFormatCheckbox: preserveNumberFormatCheckbox,
            previewCheckbox: previewCheckbox,
            eraLabel: eraLabel,
            weekdayLabel: weekdayLabel,
            daysDiffLabel: daysDiffLabel,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * 年・月・日の入力欄を整数として読む（数値でなければ NaN）
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @returns {{year: number, month: number, day: number}} 入力された日付
     */
    function readDateInputs(dialogControls) {
        return {
            year: parseInt(dialogControls.yearInput.text, 10),
            month: parseInt(dialogControls.monthInput.text, 10),
            day: parseInt(dialogControls.dayInput.text, 10)
        };
    }

    /**
     * ダイアログの入力から置換の設定を読み取る
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @returns {ReplaceOptions} 置換の設定
     */
    function readReplaceOptions(dialogControls) {
        var inputDate = readDateInputs(dialogControls);
        var weekdayChoice = getDropdownValue(dialogControls.weekdayDropdown, WEEKDAY_VALUES, 'none');
        return {
            newYear: inputDate.year,
            newMonth: inputDate.month,
            newDay: inputDate.day,
            formatChoice: getDropdownValue(dialogControls.formatDropdown, FORMAT_VALUES, 'preserve'),
            weekdayChoice: weekdayChoice,
            preserveNumberFormat: dialogControls.preserveNumberFormatCheckbox.value,
            explicitWeekdayText: buildExplicitWeekday(weekdayChoice, inputDate.year, inputDate.month, inputDate.day)
        };
    }

    /**
     * 入力値を検証する
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @param {Object[]} foundMatches - 見つかった日付
     * @returns {string|null} エラーがあればメッセージ、なければ null
     */
    function getValidationErrorMessage(dialogControls, foundMatches) {
        var inputDate = readDateInputs(dialogControls);
        var y = inputDate.year;
        var m = inputDate.month;
        var d = inputDate.day;
        if (isNaN(y) || isNaN(m) || isNaN(d)) return "年・月・日は半角数字で入力してください。";
        if (y < 1) return "年は1以上で入力してください。";
        if (m < 1 || m > 12) return "月は1〜12で入力してください。";
        if (d < 1 || d > 31) return "日は1〜31で入力してください。";
        if (!isRealDate(y, m, d)) return "存在しない日付です。";
        var checkDate = new Date(y, m - 1, d);

        /* 元号フォーマット選択時は、各元号の有効範囲に収まることを要求 */
        var formatChoice = getDropdownValue(dialogControls.formatDropdown, FORMAT_VALUES, 'preserve');
        var isReiwaFormat = (formatChoice === 'reiwa-jp' || formatChoice === 'r-dot' || formatChoice === 'r-slash');
        var isHeiseiFormat = (formatChoice === 'heisei-jp' || formatChoice === 'h-dot' || formatChoice === 'h-slash');
        if (isReiwaFormat && checkDate < REIWA_START_DATE) {
            return "令和形式は 2019/5/1 以降の日付で指定してください。";
        }
        if (isHeiseiFormat && (checkDate < HEISEI_START_DATE || checkDate >= HEISEI_END_DATE_EXCLUSIVE)) {
            return "平成形式は 1989/1/8〜2019/4/30 の範囲で指定してください。";
        }

        /* 「元の形式を保持」では、チェック済みマッチの元号がそれぞれ有効になる範囲を要求 */
        if (formatChoice === 'preserve') {
            var hasReiwaMatch = false, hasHeiseiMatch = false;
            for (var fmIdx = 0; fmIdx < foundMatches.length; fmIdx++) {
                var foundMatch = foundMatches[fmIdx];
                if (foundMatch.checkbox && !foundMatch.checkbox.value) continue;
                if (!foundMatch.parsed) continue;
                if (foundMatch.parsed.era === 'reiwa') hasReiwaMatch = true;
                if (foundMatch.parsed.era === 'heisei') hasHeiseiMatch = true;
            }
            if (hasReiwaMatch && checkDate < REIWA_START_DATE) {
                return "令和形式のマッチが含まれているため、2019/5/1 以降の日付を指定してください。";
            }
            if (hasHeiseiMatch && (checkDate < HEISEI_START_DATE || checkDate >= HEISEI_END_DATE_EXCLUSIVE)) {
                return "平成形式のマッチが含まれているため、1989/1/8〜2019/4/30 の日付を指定してください。";
            }
        }

        return null;
    }

    /**
     * 確認パネル（和暦・曜日・最初の日付との差）を入力値で更新する
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @param {Date|null} referenceDate - 比較の基準日（最初に見つかった日付）
     * @returns {void}
     */
    function updateCheckPanel(dialogControls, referenceDate) {
        var inputDate = readDateInputs(dialogControls);
        var y = inputDate.year;
        var m = inputDate.month;
        var d = inputDate.day;

        dialogControls.eraLabel.text = formatEraLabel(y, m, d);
        dialogControls.weekdayLabel.text = getWeekdayLabel(y, m, d);

        dialogControls.daysDiffLabel.text = "";
        if (referenceDate && !isNaN(y) && !isNaN(m) && !isNaN(d)) {
            var newDate = new Date(y, m - 1, d);
            if (!isNaN(newDate.getTime())) {
                dialogControls.daysDiffLabel.text = formatDaysDifference(getDaysDifference(referenceDate, newDate));
            }
        }
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビューの適用・巻き戻しを受け持つ（巻き戻しは適用した文字操作数だけ app.undo()）
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @param {Object[]} foundMatches - 見つかった日付
     * @param {TextFrame[]} textFrames - 検索対象のテキストフレーム
     * @returns {{apply: Function, revert: Function, refresh: Function}} プレビューの操作
     */
    function createPreviewController(dialogControls, foundMatches, textFrames) {
        var isPreviewApplied = false;
        var previewOpCount = 0;

        /* プレビューを巻き戻す / Undo the preview */
        function revertPreview() {
            if (!isPreviewApplied) return;
            for (var i = 0; i < previewOpCount; i++) {
                try { app.undo(); } catch (eUndo) { break; }
            }
            isPreviewApplied = false;
            previewOpCount = 0;
            app.redraw();
        }

        /* プレビューを適用（呼び出し前に入力が有効であることを確認しておく）/ Apply the preview (input must be valid) */
        function applyPreview() {
            var replaceResult = performReplacement(foundMatches, textFrames, readReplaceOptions(dialogControls));
            previewOpCount = replaceResult.operations;
            isPreviewApplied = replaceResult.operations > 0;
            app.redraw();
        }

        /* 入力やチェック変更時の自動更新（プレビュー ON 時のみ巻き戻し→再適用）/ Re-apply while the preview is on */
        function refreshPreview() {
            if (!dialogControls.previewCheckbox.value) return;
            revertPreview();
            if (getValidationErrorMessage(dialogControls, foundMatches) !== null) return;
            applyPreview();
        }

        return {
            apply: applyPreview,
            revert: revertPreview,
            refresh: refreshPreview
        };
    }

    /**
     * ダイアログにイベントを付け、確認パネルなどの初期状態を反映する
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @param {Object[]} foundMatches - 見つかった日付
     * @param {Object} previewController - createPreviewController() の戻り値
     * @returns {void}
     */
    function bindDialogEvents(dialogControls, foundMatches, previewController) {
        var dateDialog = dialogControls.dialog;
        var dateCheckboxes = dialogControls.dateCheckboxes;

        /* 比較用の基準日（最初に見つかった日付。形式に関わらず西暦で扱う） */
        var referenceDate = null;
        if (foundMatches.length > 0 && foundMatches[0].parsed) {
            var referenceParsed = foundMatches[0].parsed;
            referenceDate = new Date(referenceParsed.year, referenceParsed.month - 1, referenceParsed.day);
        }

        /* チェックボックスに Option＋クリックで全切替を割り当てる。クリック後はプレビューも更新 */
        function bindOptionClickToggleAll(checkboxControl) {
            checkboxControl.addEventListener("click", function (event) {
                if (event.altKey) {
                    var newValue = checkboxControl.value;
                    for (var idx = 0; idx < dateCheckboxes.length; idx++) {
                        dateCheckboxes[idx].value = newValue;
                    }
                }
                previewController.refresh();
            });
        }
        for (var i = 0; i < dateCheckboxes.length; i++) {
            bindOptionClickToggleAll(dateCheckboxes[i]);
        }

        /* 入力値からチェック用の表示を更新。プレビュー ON 時は再適用も走らせる */
        function refreshCheckLabels() {
            updateCheckPanel(dialogControls, referenceDate);
            previewController.refresh();
        }

        var dateInputs = [dialogControls.yearInput, dialogControls.monthInput, dialogControls.dayInput];
        for (var j = 0; j < dateInputs.length; j++) {
            dateInputs[j].onChange = refreshCheckLabels;
        }
        for (var k = 0; k < dateInputs.length; k++) {
            changeValueByArrowKey(dateInputs[k], refreshCheckLabels);
        }

        refreshCheckLabels();

        /* 「数字の書式を保持」は preserve 選択時のみ有効。フォーマット切替で enabled を連動させる */
        function updatePreserveNumberFormatEnabled() {
            var formatChoice = getDropdownValue(dialogControls.formatDropdown, FORMAT_VALUES, 'preserve');
            dialogControls.preserveNumberFormatCheckbox.enabled = (formatChoice === 'preserve');
        }
        updatePreserveNumberFormatEnabled();

        /* フォーマット選択 / 曜日 / 数字書式保持の変更でもプレビューを更新 */
        dialogControls.formatDropdown.onChange = function () {
            updatePreserveNumberFormatEnabled();
            previewController.refresh();
        };
        dialogControls.weekdayDropdown.onChange = function () { previewController.refresh(); };
        dialogControls.preserveNumberFormatCheckbox.onClick = function () { previewController.refresh(); };

        /* プレビューチェックボックス本体 */
        dialogControls.previewCheckbox.onClick = function () {
            if (dialogControls.previewCheckbox.value) {
                var validationError = getValidationErrorMessage(dialogControls, foundMatches);
                if (validationError) {
                    alert(validationError);
                    dialogControls.previewCheckbox.value = false;
                    return;
                }
                previewController.apply();
            } else {
                previewController.revert();
            }
        };

        dialogControls.btnOK.onClick = function () {
            var validationError = getValidationErrorMessage(dialogControls, foundMatches);
            if (validationError) {
                alert(validationError);
                return;
            }
            dateDialog.close(1);
        };

        dialogControls.btnCancel.onClick = function () {
            previewController.revert();
            dateDialog.close(0);
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 日付を探してダイアログを表示し、チェックした日付を置換する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert("ドキュメントが開かれていません。");
            return;
        }

        var doc = app.activeDocument;
        var textFrames = collectTargetTextFrames(doc);
        var foundMatches = findDateMatches(doc, textFrames);
        pairWeekdayFrames(foundMatches, textFrames);

        var dialogControls = buildDialog(doc, foundMatches);
        var previewController = createPreviewController(dialogControls, foundMatches, textFrames);
        bindDialogEvents(dialogControls, foundMatches, previewController);

        if (dialogControls.dialog.show() !== 1) {
            return;
        }

        if (foundMatches.length === 0) {
            alert("置換対象の日付がありません。");
            return;
        }

        /* OK時はプレビュー結果をそのまま確定せず、いったん戻してから本番置換を再実行する */
        previewController.revert();

        var replaceResult = performReplacement(foundMatches, textFrames, readReplaceOptions(dialogControls));

        if (replaceResult.selectedCount === 0) {
            alert("置換する項目が選択されていません。");
            return;
        }

        if (replaceResult.errors > 0) {
            alert(formatErrorMessage(replaceResult.errors));
        }
    }

    main();

})();
