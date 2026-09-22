#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストフレーム内の数字・英字・日付・時刻を検出し、値を増分しながら下方向へ複製します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartIncrementText.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n5f25ed17b123

### Overview

Finds the digits, letters, dates or times in the selected text frame and duplicates it downwards,
incrementing the value each time.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartIncrementText.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartIncrementText";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartIncrementText.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartIncrementText.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n5f25ed17b123"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    var DEFAULT_COPY_COUNT = 5;        /* ［複製数］の初期値 / initial number of copies */
    var DEFAULT_STEP = 1;              /* ［増分］の初期値 / initial step */
    var DEFAULT_PITCH_RATIO = 1.5;     /* 既定の送り＝文字サイズ×この倍率 / default pitch = font size * this */
    var DEFAULT_ZERO_PAD = true;       /* ［ゼロ埋め］の初期状態 / zero padding on by default */
    var DEFAULT_MERGE_ON_OK = false;   /* ［確定時にテキストを結合］の初期状態 / merging off by default */
    var FALLBACK_FONT_SIZE_PT = 10;    /* 文字サイズを取得できないときの代替値 / fallback font size */

    /* ダイアログ位置を覚えておく環境設定キー / Preference key that remembers the dialog position */
    var DIALOG_POSITION_PREF_KEY = "dupTextWithIncrementNumbers_v2_dialog_pos";

    // =========================================
    // レイアウト / Layout
    // =========================================
    var FIELD_LABEL_WIDTH = 60;        /* 項目名の幅 / width of a row label */
    var NUMBER_FIELD_CHARS = 4;        /* 数値入力欄の文字数 / width of a numeric field */
    var START_FIELD_CHARS = 6;         /* ［開始値］入力欄の文字数 / width of the start field */
    var BUTTON_ROW_TOP_MARGIN = 10;    /* ボタンエリアの上余白 / top margin of the button row */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UIの表示言語を返す
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "連番複製", en: "Duplicate with Increment" }
        },
        fieldLabel: {
            copyCount: { ja: "複製数", en: "Copies" },
            stepValue: { ja: "増分", en: "Step" },
            gap: { ja: "アキ", en: "Gap" },
            incrementTarget: { ja: "増分対象", en: "Target" }
        },
        checkbox: {
            startOverride: { ja: "開始値", en: "Start value" },
            zeroPad: { ja: "ゼロ埋め", en: "Zero pad" },
            mergeOnOK: { ja: "1つのテキストに結合", en: "Merge into one text" }
        },
        radio: {
            year: { ja: "年", en: "Year" },
            month: { ja: "月", en: "Month" },
            day: { ja: "日", en: "Day" },
            hour: { ja: "時", en: "Hour" },
            minute: { ja: "分", en: "Minute" },
            numberPrefix: { ja: "数字", en: "Num" },
            alphabetPrefix: { ja: "英字", en: "Alpha" }
        },
        tooltip: {
            copyCount: { ja: "作る複製の数です。元のテキストは含みません。", en: "How many copies to create, not counting the original." },
            stepValue: { ja: "1つ進むごとに足す数です。負の値で減らせます。", en: "Amount added at each step. Negative values count down." },
            gap: { ja: "複製どうしのアキです。文字サイズに加算されます。", en: "Space between copies, added on top of the font size." },
            incrementTarget: {
                ja: "テキストの中で増やす箇所です。数字や英字が複数あるときに選べます。",
                en: "Which part of the text to increment, when there is more than one number or letter."
            },
            startOverride: { ja: "元のテキストの値ではなく、指定した値から始めます。", en: "Starts from the value you enter instead of the one in the original text." },
            startValue: { ja: "最初の複製に使う値です。", en: "Value used for the first copy." },
            zeroPad: {
                ja: "元の桁数に合わせて、頭に0を足します。最終値で桁が増えるときは桁数を広げます。",
                en: "Pads with leading zeros to match the original width, widening it when the last value needs another digit."
            },
            mergeOnOK: {
                ja: "OKを押したとき、複製したテキストを1つのテキストにまとめます。",
                en: "Merges the duplicated text into a single text object when you press OK."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectOneTextFrame: { ja: "テキストフレームを1つだけ選択してください。", en: "Select exactly one text frame." },
            noTarget: {
                ja: "増分対象が見つかりませんでした。\n英字は1文字（A〜Z）の場合のみ対応します。\n例: A1 はOK / AB1, Ver1 は英字増分対象になりません。",
                en: "No increment target was found.\nAlphabet increment supports only single-letter tokens (A–Z).\nExamples: A1 is OK / AB1, Ver1 are NOT alphabet targets."
            },
            noToken: {
                ja: "選択されたテキストの中に「半角数字」または「英字」が見つかりませんでした。\n数字/英字を含むテキスト（例: 01, 2025/11/21, 19:00, A1）を選択してください。",
                en: "No digits or letters were found in the selected text.\nSelect text containing digits/letters (e.g., 01, 2025/11/21, 19:00, A1)."
            }
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

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {object} labelSet - ラベル定義
     * @returns {string} コロンを付けた項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ": ");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * ポイント値を現在の単位の値に換算する
     * @param {number} pointValue - ポイント値
     * @param {{pointsPerUnit: number}} unitInfo - 単位の情報
     * @returns {number} 換算後の値
     */
    function ptToUnitValue(pointValue, unitInfo) {
        return pointValue / unitInfo.pointsPerUnit;
    }

    /**
     * 単位付きの値をポイント値に換算する
     * @param {number} unitValue - 単位の値
     * @param {{pointsPerUnit: number}} unitInfo - 単位の情報
     * @returns {number} ポイント値
     */
    function unitValueToPt(unitValue, unitInfo) {
        return unitValue * unitInfo.pointsPerUnit;
    }

    var textUnitInfo = getUnitInfo("text/units");

    // =========================================
    // 事前チェック / Pre-check
    // =========================================
    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.noDocument));
        return;
    }

    var doc = app.activeDocument;
    var selectedItems = doc.selection;
    if (selectedItems.length !== 1 || selectedItems[0].typename !== "TextFrame") {
        alert(getLabel(LABELS.alert.selectOneTextFrame));
        return;
    }

    var sourceTextFrame = selectedItems[0];
    var sourceText = sourceTextFrame.contents; /* 復帰用にも使う / also used to restore */

    // =========================================
    // 元テキストの解析 / Analyze the source text
    // =========================================

    /* 数字・英字のまとまりを抽出し、リテラル部分とトークンに分ける
       例: "A1" -> ["A", "1"] / "AB1" -> ["AB", "1"]（AB は英字増分の対象外）
       Split the text into literal segments and digit/letter tokens */
    var literalSegments = [];
    var sourceTokens = [];
    var tokenTypes = []; /* "alpha1"（英字1文字）/ "num"（数字）/ "other" */
    var tokenPattern = /[A-Za-z]+|\d+/g;
    var tokenMatch;
    var literalStart = 0;

    while ((tokenMatch = tokenPattern.exec(sourceText)) !== null) {
        literalSegments.push(sourceText.substring(literalStart, tokenMatch.index));
        sourceTokens.push(tokenMatch[0]);
        tokenTypes.push(/^[A-Za-z]$/.test(tokenMatch[0]) ? "alpha1" : (/^\d+$/.test(tokenMatch[0]) ? "num" : "other"));
        literalStart = tokenMatch.index + tokenMatch[0].length;
    }
    literalSegments.push(sourceText.substring(literalStart));

    if (sourceTokens.length === 0) {
        alert(getLabel(LABELS.alert.noToken));
        return;
    }

    /**
     * トークン配列からテキストを組み立て直す
     * @param {string[]} tokens - 差し替え後のトークン
     * @returns {string} 組み立てたテキスト
     */
    function rebuildText(tokens) {
        var rebuiltText = literalSegments[0];
        for (var i = 0; i < tokens.length; i++) {
            rebuiltText += String(tokens[i]) + literalSegments[i + 1];
        }
        return rebuiltText;
    }

    /**
     * 増分対象のラジオに表示する名前を返す
     * @param {string} kind - "date_ymd" / "time_hm" / "alpha1" / "num"
     * @param {number} position - 同じ種類の中での位置（1始まり）
     * @returns {string} ラジオのラベル
     */
    function getTargetRadioLabel(kind, position) {
        if (kind === "date_ymd") {
            if (position === 1) return getLabel(LABELS.radio.year);
            if (position === 2) return getLabel(LABELS.radio.month);
            return getLabel(LABELS.radio.day);
        }
        if (kind === "time_hm") {
            return getLabel((position === 1) ? LABELS.radio.hour : LABELS.radio.minute);
        }
        if (kind === "alpha1") return getLabel(LABELS.radio.alphabetPrefix) + position;
        return getLabel(LABELS.radio.numberPrefix) + position;
    }

    /**
     * 元テキストの書式（日付／時刻／一般）を判定し、増分対象の候補を組み立てる
     * @returns {{patternType: string, radioLabels: string[], tokenIndices: number[]}} 判定結果
     */
    function detectIncrementTargets() {
        /* 和文年月日 / Japanese YMD (with optional weekday) */
        var dateJaPattern = /^(\d{4})年(\d{1,2})月(\d{1,2})日(?:[（(［\[]?(?:日|月|火|水|木|金|土)[）)\]］]?)?$/;
        /* スラッシュ区切り / Slash YMD (with optional weekday) */
        var dateSlashPattern = /^(\d{4})\/(\d{1,2})\/(\d{1,2})(?:[ 　\t]*[（(［\[]?(?:日|月|火|水|木|金|土|Sun|Mon|Tue|Wed|Thu|Fri|Sat)[）)\]］]?)?$/;
        /* ドット区切り3要素 / Dot 3 parts */
        var dateDotPattern = /^(\d+)\.(\d+)\.(\d+)$/;
        /* 時刻 / Time */
        var timePattern = /^(\d{1,2}):(\d{2})$/;

        if (dateJaPattern.test(sourceText) || dateSlashPattern.test(sourceText) || dateDotPattern.test(sourceText)) {
            return {
                patternType: "date_ymd",
                radioLabels: [getTargetRadioLabel("date_ymd", 1), getTargetRadioLabel("date_ymd", 2), getTargetRadioLabel("date_ymd", 3)],
                tokenIndices: [0, 1, 2]
            };
        }
        if (timePattern.test(sourceText)) {
            return {
                patternType: "time_hm",
                radioLabels: [getTargetRadioLabel("time_hm", 1), getTargetRadioLabel("time_hm", 2)],
                tokenIndices: [0, 1]
            };
        }

        var radioLabels = [];
        var tokenIndices = [];
        var numberCount = 0;
        var alphabetCount = 0;
        for (var i = 0; i < sourceTokens.length; i++) {
            /* "AB" や "Ver" のような2文字以上の英字は候補にしない / multi-letter tokens are not candidates */
            if (tokenTypes[i] === "alpha1") {
                alphabetCount++;
                radioLabels.push(getTargetRadioLabel("alpha1", alphabetCount));
                tokenIndices.push(i);
            } else if (tokenTypes[i] === "num") {
                numberCount++;
                radioLabels.push(getTargetRadioLabel("num", numberCount));
                tokenIndices.push(i);
            }
        }
        return { patternType: "generic", radioLabels: radioLabels, tokenIndices: tokenIndices };
    }

    var incrementTargets = detectIncrementTargets();
    var patternType = incrementTargets.patternType;
    var targetRadioLabels = incrementTargets.radioLabels;
    var targetTokenIndices = incrementTargets.tokenIndices;

    if (targetTokenIndices.length === 0) {
        alert(getLabel(LABELS.alert.noTarget));
        return;
    }

    /* 既定の増分対象：日付は日、時刻は分、一般は末尾の候補
       Default target: day for dates, minute for times, the last candidate otherwise */
    var targetTokenIndex = targetTokenIndices[targetTokenIndices.length - 1];
    if (patternType === "date_ymd") targetTokenIndex = 2;
    else if (patternType === "time_hm") targetTokenIndex = 1;

    var targetDigitLength = String(sourceTokens[targetTokenIndex]).length;

    /* 複製の送りと結合時の行送りの基準にする文字サイズ / font size used for the pitch and the merged leading */
    var sourceFontSizePt = sourceTextFrame.textRange.characterAttributes.size;
    if (isNaN(sourceFontSizePt) || sourceFontSizePt <= 0) sourceFontSizePt = FALLBACK_FONT_SIZE_PT;

    // =========================================
    // 増分の計算 / Increment helpers
    // =========================================

    /**
     * 先頭に0を足して桁を揃える（負の値は符号を残して数字だけを揃える）
     * @param {string|number} value - 対象の値
     * @param {number} length - 揃える桁数
     * @returns {string} ゼロ埋めした文字列
     */
    function zeroPad(value, length) {
        var digits = String(value);
        var sign = "";
        if (digits.charAt(0) === "-") {
            sign = "-";
            digits = digits.substring(1);
        }
        while (digits.length < length) digits = "0" + digits;
        return sign + digits;
    }

    /**
     * 英字1文字のトークンかを判定する
     * @param {string} token - 判定するトークン
     * @returns {boolean} A〜Z／a〜z 1文字なら true
     */
    function isAlphaToken(token) {
        return /^[A-Za-z]$/.test(String(token));
    }

    /**
     * 英字1文字を番号に変換する（A=1〜Z=26）
     * @param {string} alphaToken - 英字1文字
     * @returns {number|null} 番号。英字1文字でなければ null
     */
    function alphaToNumber(alphaToken) {
        var upperCased = String(alphaToken).toUpperCase();
        if (!/^[A-Z]$/.test(upperCased)) return null;
        return upperCased.charCodeAt(0) - 64;
    }

    /**
     * 番号を英字1文字に変換する（26を超えたらAに戻る）
     * @param {number} value - 番号
     * @param {boolean} isLowerCase - 小文字で返すか
     * @returns {string} 英字1文字
     */
    function numberToAlpha(value, isLowerCase) {
        var letterNumber = Math.floor(Number(value));
        if (!isFinite(letterNumber) || letterNumber <= 0) letterNumber = 1;
        var offset = ((letterNumber - 1) % 26 + 26) % 26;
        var alphaToken = String.fromCharCode(65 + offset);
        return isLowerCase ? alphaToken.toLowerCase() : alphaToken;
    }

    /**
     * トークンが小文字かを判定する
     * @param {string} token - 判定するトークン
     * @returns {boolean} 小文字なら true
     */
    function isLowerCaseToken(token) {
        return String(token) === String(token).toLowerCase();
    }

    /**
     * テキストに含まれる曜日表記を調べる
     * @param {string} text - 調べるテキスト
     * @returns {{has: boolean, style: string, left: string, right: string}} 曜日の有無と囲み記号
     */
    function detectWeekdayInfo(text) {
        var weekdayMatch = text.match(/([（(［\[])[ 　\t]*(日|月|火|水|木|金|土|Sun|Mon|Tue|Wed|Thu|Fri|Sat)[ 　\t]*([）)\]］])/);
        if (!weekdayMatch) return { has: false };
        return { has: true, style: (weekdayMatch[2].length === 1) ? "ja" : "en", left: weekdayMatch[1], right: weekdayMatch[3] };
    }

    var weekdayInfo = detectWeekdayInfo(sourceText);

    /**
     * 日付に対応する曜日の文字を返す
     * @param {Date} date - 対象の日付
     * @param {string} style - "ja"（日〜土）または "en"（Sun〜Sat）
     * @returns {string} 曜日の文字
     */
    function weekdayTokenByDate(date, style) {
        var weekdaysJa = ["日", "月", "火", "水", "木", "金", "土"];
        var weekdaysEn = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
        return (style === "en") ? weekdaysEn[date.getDay()] : weekdaysJa[date.getDay()];
    }

    /**
     * テキスト内の曜日表記を日付に合わせて書き換える
     * @param {string} text - 対象のテキスト
     * @param {Date} date - 合わせる日付
     * @returns {string} 書き換えたテキスト
     */
    function applyWeekdayToText(text, date) {
        if (!weekdayInfo.has) return text;
        var weekdayPattern = /([（(［\[])[ 　\t]*(日|月|火|水|木|金|土|Sun|Mon|Tue|Wed|Thu|Fri|Sat)[ 　\t]*([）)\]］])/;
        return text.replace(weekdayPattern, weekdayInfo.left + weekdayTokenByDate(date, weekdayInfo.style) + weekdayInfo.right);
    }

    /**
     * 年月日のトークンから日付を作る
     * @param {string[]} tokens - トークン配列
     * @returns {Date|null} 日付。数値として読めなければ null
     */
    function parseDateFromTokens(tokens) {
        if (tokens.length < 3) return null;
        var year = parseInt(tokens[0], 10);
        var month = parseInt(tokens[1], 10);
        var day = parseInt(tokens[2], 10);
        if (isNaN(year) || isNaN(month) || isNaN(day)) return null;

        /* 2桁年などは現在の世紀で補う / short years are completed with the current century */
        if (String(sourceTokens[0]).length < 4) {
            year = Math.floor((new Date()).getFullYear() / 100) * 100 + year;
        }
        return new Date(year, month - 1, day);
    }

    /**
     * 日付を元テキストの桁数に合わせたトークンにする
     * @param {Date} date - 対象の日付
     * @returns {string[]} 年・月・日のトークン
     */
    function formatDateToTokens(date) {
        var yearLength = String(sourceTokens[0]).length;
        var monthLength = String(sourceTokens[1]).length;
        var dayLength = String(sourceTokens[2]).length;

        var yearToken = String(date.getFullYear());
        if (yearLength < yearToken.length) yearToken = yearToken.slice(yearToken.length - yearLength);
        if (yearLength > yearToken.length) yearToken = zeroPad(yearToken, yearLength);

        var monthToken = String(date.getMonth() + 1);
        var dayToken = String(date.getDate());
        if (monthLength > 1) monthToken = zeroPad(monthToken, monthLength);
        if (dayLength > 1) dayToken = zeroPad(dayToken, dayLength);

        return [yearToken, monthToken, dayToken];
    }

    /**
     * 日付を年／月／日の単位で増減する
     * @param {Date} baseDate - 基準の日付
     * @param {number} unitIndex - 0=年 / 1=月 / 2=日
     * @param {number} step - 増減する量
     * @returns {Date} 増減後の日付
     */
    function addDateByUnit(baseDate, unitIndex, step) {
        var date = new Date(baseDate.getTime());
        if (unitIndex === 0) date.setFullYear(date.getFullYear() + step);
        else if (unitIndex === 1) date.setMonth(date.getMonth() + step);
        else date.setDate(date.getDate() + step);
        return date;
    }

    /**
     * 時分のトークンから時刻を作る
     * @param {string[]} tokens - トークン配列
     * @returns {{hour: number, minute: number}|null} 時刻。数値として読めなければ null
     */
    function parseTimeFromTokens(tokens) {
        if (tokens.length < 2) return null;
        var hour = parseInt(tokens[0], 10);
        var minute = parseInt(tokens[1], 10);
        if (isNaN(hour) || isNaN(minute)) return null;
        return { hour: hour, minute: minute };
    }

    /**
     * 時刻を時／分の単位で増減する（24時間で繰り上がる）
     * @param {{hour: number, minute: number}} baseTime - 基準の時刻
     * @param {number} unitIndex - 0=時 / 1=分
     * @param {number} step - 増減する量
     * @returns {{hour: number, minute: number}} 増減後の時刻
     */
    function addTimeByUnit(baseTime, unitIndex, step) {
        var totalMinutes = baseTime.hour * 60 + baseTime.minute + ((unitIndex === 0) ? step * 60 : step);
        totalMinutes = ((totalMinutes % 1440) + 1440) % 1440;
        return { hour: Math.floor(totalMinutes / 60), minute: totalMinutes % 60 };
    }

    /**
     * 時刻を元テキストの桁数に合わせたトークンにする
     * @param {{hour: number, minute: number}} time - 対象の時刻
     * @returns {string[]} 時・分のトークン
     */
    function formatTimeToTokens(time) {
        var hourLength = String(sourceTokens[0]).length;
        var minuteLength = String(sourceTokens[1]).length;
        var hourToken = String(time.hour);
        var minuteToken = String(time.minute);
        if (hourLength > 1) hourToken = zeroPad(hourToken, hourLength);
        if (minuteLength > 1) minuteToken = zeroPad(minuteToken, minuteLength);
        return [hourToken, minuteToken];
    }

    // =========================================
    // UIの値の取り出し / Reading the dialog values
    // =========================================

    /**
     * ［増分］の値を取り出す（0は1として扱う）
     * @returns {number} 増分の値
     */
    function getStepValue() {
        var stepValue = parseInt(String(stepInput.text), 10);
        return (isNaN(stepValue) || stepValue === 0) ? 1 : stepValue;
    }

    /**
     * ［開始値］に有効な値が入っているかを判定する
     * @returns {boolean} 使える値なら true
     */
    function isStartValueValid() {
        if (!startOverrideCheckbox.value) return true;
        var rawValue = String(startValueInput.text);
        if (tokenTypes[targetTokenIndex] === "alpha1") return isAlphaToken(rawValue);
        return !isNaN(parseInt(rawValue, 10));
    }

    /**
     * ［開始値］を反映した、増分前のトークン配列を作る
     * @returns {string[]} トークン配列
     */
    function buildBaseTokens() {
        var tokens = sourceTokens.slice();
        if (!startOverrideCheckbox.value) return tokens;

        var rawValue = String(startValueInput.text);
        if (tokenTypes[targetTokenIndex] === "alpha1") {
            if (isAlphaToken(rawValue)) tokens[targetTokenIndex] = rawValue;
        } else {
            var startNumber = parseInt(rawValue, 10);
            if (!isNaN(startNumber)) tokens[targetTokenIndex] = String(startNumber);
        }
        return tokens;
    }

    /**
     * ［ゼロ埋め］ON時の桁数を返す（最終値が桁上がりするときは広げる）
     * @param {number} baseNumber - 開始の値
     * @param {number} baseLength - 元の桁数
     * @returns {number} 使用する桁数
     */
    function getPadLength(baseNumber, baseLength) {
        var copyCount = parseInt(countInput.text, 10);
        if (isNaN(copyCount) || copyCount < 0) return baseLength;
        var lastNumber = baseNumber + (copyCount * getStepValue());
        /* 桁数は符号を除いて数える / the minus sign does not count as a digit */
        var neededLength = Math.max(String(Math.abs(baseNumber)).length, String(Math.abs(lastNumber)).length);
        return (neededLength > baseLength) ? neededLength : baseLength;
    }

    /**
     * ［ゼロ埋め］の設定に従って数値を文字列にする
     * @param {number} value - 変換する値
     * @param {number} baseNumber - 開始の値
     * @param {number} baseLength - 元の桁数
     * @returns {string} 文字列にした値
     */
    function formatNumberToken(value, baseNumber, baseLength) {
        if (!zeroPadCheckbox.value) return String(value);
        return zeroPad(value, getPadLength(baseNumber, baseLength));
    }

    // =========================================
    // テキストの生成 / Building the text
    // =========================================

    /**
     * 対象トークンを数値として増分したテキストを作る
     * @param {string[]} tokens - 元にするトークン配列（書き換える）
     * @param {number} offset - 開始の値からの増分
     * @param {boolean} usePadding - ［ゼロ埋め］の設定を反映するか
     * @returns {string} 組み立てたテキスト
     */
    function buildNumberIncrementedText(tokens, offset, usePadding) {
        var baseNumber = parseInt(tokens[targetTokenIndex], 10);
        if (isNaN(baseNumber)) baseNumber = 0;
        var value = baseNumber + offset;
        tokens[targetTokenIndex] = usePadding ? formatNumberToken(value, baseNumber, targetDigitLength) : String(value);
        return rebuildText(tokens);
    }

    /**
     * 複製1つ分のテキストを作る
     * @param {string[]} baseTokens - 増分前のトークン配列
     * @param {number} offset - 開始の値からの増分（0で複製元と同じ値）
     * @returns {string} 複製に入れるテキスト
     */
    function buildIncrementedText(baseTokens, offset) {
        var tokens = baseTokens.slice();

        if (patternType === "date_ymd") {
            var baseDate = parseDateFromTokens(baseTokens);
            /* 暦として読めない並びは、対象トークンを数値として増分する / fall back to a plain number */
            if (!baseDate) return buildNumberIncrementedText(tokens, offset, false);
            var shiftedDate = addDateByUnit(baseDate, targetTokenIndex, offset);
            var dateTokens = formatDateToTokens(shiftedDate);
            tokens[0] = dateTokens[0];
            tokens[1] = dateTokens[1];
            tokens[2] = dateTokens[2];
            return applyWeekdayToText(rebuildText(tokens), shiftedDate);
        }

        if (patternType === "time_hm") {
            var baseTime = parseTimeFromTokens(baseTokens);
            if (!baseTime) return buildNumberIncrementedText(tokens, offset, false);
            var timeTokens = formatTimeToTokens(addTimeByUnit(baseTime, targetTokenIndex, offset));
            tokens[0] = timeTokens[0];
            tokens[1] = timeTokens[1];
            return rebuildText(tokens);
        }

        if (tokenTypes[targetTokenIndex] === "alpha1") {
            var baseAlphaToken = String(baseTokens[targetTokenIndex]);
            var alphaNumber = alphaToNumber(baseAlphaToken);
            if (alphaNumber == null) alphaNumber = 1;
            tokens[targetTokenIndex] = numberToAlpha(alphaNumber + offset, isLowerCaseToken(baseAlphaToken));
            return rebuildText(tokens);
        }

        return buildNumberIncrementedText(tokens, offset, true);
    }

    /**
     * 現在のUI設定を反映した複製元のテキストを作る
     * @returns {string|null} 反映後のテキスト。［開始値］が不正なときは null
     */
    function buildSourceContents() {
        if (!isStartValueValid()) return null;

        /* 数字が対象で、ゼロ埋めも開始値の指定もないときは元の表記のまま
           Keep the original spelling when nothing reformats the number */
        if (patternType === "generic" && tokenTypes[targetTokenIndex] !== "alpha1" &&
            !zeroPadCheckbox.value && !startOverrideCheckbox.value) {
            return sourceText;
        }
        return buildIncrementedText(buildBaseTokens(), 0);
    }

    // =========================================
    // ドキュメントの更新 / Updating the document
    // =========================================

    /**
     * テキストフレームの位置を控える（contents の書き換えで位置がずれるため）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {number[]} [x, y]
     */
    function getTextFramePosition(textFrame) {
        return [textFrame.position[0], textFrame.position[1]];
    }

    /**
     * 控えておいた位置にテキストフレームを戻す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number[]} position - [x, y]
     * @returns {void}
     */
    function setTextFramePosition(textFrame, position) {
        textFrame.position = [position[0], position[1]];
    }

    /**
     * 複製元のテキストを、位置を保ったまま書き換える
     * @param {string} text - 設定するテキスト
     * @returns {void}
     */
    function setSourceContents(text) {
        var savedPosition = getTextFramePosition(sourceTextFrame);
        sourceTextFrame.contents = text;
        setTextFramePosition(sourceTextFrame, savedPosition);
    }

    /**
     * 現在のUI設定を複製元のテキストに反映する
     * @returns {void}
     */
    function applySettingsToSourceText() {
        var sourceContents = buildSourceContents();
        if (sourceContents === null) return;
        setSourceContents(sourceContents);
    }

    /**
     * 指定した数だけ複製を作り、値を増分する
     * @returns {TextFrame[]} 作成した複製
     */
    function createIncrementedCopies() {
        var copyCount = parseInt(countInput.text, 10);
        var gapValue = parseFloat(gapInput.text);
        if (isNaN(copyCount) || isNaN(gapValue)) return [];

        /* ［アキ］は文字サイズに足すアキ → 実際の送り＝文字サイズ＋アキ
              The gap field holds the space added on top of the font size */
        var pitchPt = sourceFontSizePt + unitValueToPt(gapValue, textUnitInfo);
        var stepValue = getStepValue();
        var baseTokens = buildBaseTokens();
        var createdCopies = [];

        for (var i = 1; i <= copyCount; i++) {
            var textCopy = sourceTextFrame.duplicate();
            textCopy.top = sourceTextFrame.top - (pitchPt * i);
            textCopy.left = sourceTextFrame.left;
            textCopy.contents = buildIncrementedText(baseTokens, i * stepValue);
            createdCopies.push(textCopy);
        }
        return createdCopies;
    }

    /**
     * 複製を改行でつないで複製元の1つのテキストにまとめる
     * @returns {void}
     */
    function mergeCopiesIntoSource() {
        var savedPosition = getTextFramePosition(sourceTextFrame);
        var mergedLines = [String(sourceTextFrame.contents)];
        for (var i = 0; i < previewItems.length; i++) {
            mergedLines.push(String(previewItems[i].contents));
        }
        sourceTextFrame.contents = mergedLines.join("\r");

        /* 行送り＝文字サイズ＋［アキ］ / Leading = font size + the gap field */
        var fontSizePt = sourceTextFrame.textRange.characterAttributes.size;
        if (isNaN(fontSizePt) || fontSizePt <= 0) fontSizePt = sourceFontSizePt;
        var gapValue = parseFloat(gapInput.text);
        if (isNaN(gapValue)) gapValue = 0;
        var leadingPt = fontSizePt + unitValueToPt(gapValue, textUnitInfo);
        if (leadingPt > 0) {
            sourceTextFrame.textRange.characterAttributes.autoLeading = false;
            sourceTextFrame.textRange.characterAttributes.leading = leadingPt;
        }

        clearPreview();
        setTextFramePosition(sourceTextFrame, savedPosition);
    }

    /**
     * 複製元のテキストと位置を実行前の状態に戻す
     * @returns {void}
     */
    function restoreSourceText() {
        previewSuspended = true;
        clearPreview();
        setSourceContents(sourceText);
        previewSuspended = false;
    }

    // =========================================
    // プレビュー / Preview
    // =========================================
    var previewItems = [];
    var previewSuspended = false;
    var closedWithOK = false;

    /**
     * プレビューで作った複製を取り除く
     * @returns {void}
     */
    function clearPreview() {
        for (var i = previewItems.length - 1; i >= 0; i--) {
            previewItems[i].remove();
        }
        previewItems = [];
    }

    /**
     * 現在のUI設定でプレビューを作り直す
     * @returns {void}
     */
    function updatePreview() {
        if (previewSuspended) return;
        applySettingsToSourceText();
        clearPreview();
        /* ［開始値］が不正な間はプレビューを出さず、［OK］も押せなくする
           No preview and no OK while the start value is invalid */
        var startValueValid = isStartValueValid();
        btnOK.enabled = startValueValid;
        if (startValueValid) previewItems = createIncrementedCopies();
        app.redraw();
    }

    // =========================================
    // ダイアログの位置 / Dialog position
    // =========================================

    /**
     * 前回のダイアログ位置を読む
     * @returns {number[]|null} [x, y]。保存されていなければ null
     */
    function readDialogPosition() {
        var savedValue = "";
        /* 未登録のキーは例外になることがある / an unset key can throw */
        try {
            savedValue = app.preferences.getStringPreference(DIALOG_POSITION_PREF_KEY);
        } catch (e) {
            return null;
        }
        if (!savedValue) return null;

        var positionParts = savedValue.split(",");
        if (positionParts.length !== 2) return null;
        var savedX = parseFloat(positionParts[0]);
        var savedY = parseFloat(positionParts[1]);
        if (isNaN(savedX) || isNaN(savedY)) return null;
        return [savedX, savedY];
    }

    /**
     * ダイアログの位置を保存する
     * @param {Window} dialogWindow - 対象のダイアログ
     * @returns {void}
     */
    function saveDialogPosition(dialogWindow) {
        var dialogLocation = dialogWindow.location;
        if (!dialogLocation) return;
        app.preferences.setStringPreference(DIALOG_POSITION_PREF_KEY, String(dialogLocation[0]) + "," + String(dialogLocation[1]));
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ↑↓キーで数値を増減できるようにする（Shiftで10、Optionで0.1）
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - 負の値を許すか
     * @param {function} onChanged - 値が変わったときに呼ぶ処理
     * @param {boolean} forceInteger - 整数に丸めるか
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative, onChanged, forceInteger) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            if (keyboard.shiftKey) {
                if (event.keyName === "Up") value = Math.ceil((value + 1) / 10) * 10;
                else value = Math.floor((value - 1) / 10) * 10;
            } else if (keyboard.altKey) {
                value += (event.keyName === "Up") ? 0.1 : -0.1;
            } else {
                value += (event.keyName === "Up") ? 1 : -1;
            }

            if (!forceInteger && keyboard.altKey) value = Math.round(value * 10) / 10;
            else value = Math.round(value);
            if (!allowNegative && value < 0) value = 0;

            event.preventDefault();
            editText.text = value;
            onChanged();
        });
    }

    /**
     * 行の左端に置く項目名を作る
     * @param {Group} targetRow - 追加先の行
     * @param {object} labelSet - 項目名のラベル定義
     * @returns {StaticText} 作成した項目名
     */
    function addRowLabel(targetRow, labelSet) {
        var rowLabel = targetRow.add("statictext", undefined, labelText(labelSet));
        rowLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        rowLabel.justify = "right";
        return rowLabel;
    }

    /**
     * 横並びの行（左寄せ・上下中央）を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @returns {Group} 作成した行
     */
    function addRow(parentGroup) {
        var newRow = parentGroup.add("group");
        newRow.orientation = "row";
        newRow.alignChildren = ["left", "center"];
        return newRow;
    }

    /**
     * 「項目名＋入力欄」の行を作る
     * @param {Group} parentGroup - 追加先のグループ
     * @param {object} labelSet - 項目名のラベル定義
     * @param {string} initialValue - 入力欄の初期値
     * @param {number} fieldChars - 入力欄の文字数
     * @param {object} tooltipSet - 入力欄のツールチップ定義
     * @returns {{fieldRow: Group, fieldInput: EditText}} 作成した行と入力欄
     */
    function addFieldRow(parentGroup, labelSet, initialValue, fieldChars, tooltipSet) {
        var fieldRow = addRow(parentGroup);
        addRowLabel(fieldRow, labelSet);

        var fieldInput = fieldRow.add("edittext", undefined, initialValue);
        fieldInput.characters = fieldChars;
        fieldInput.helpTip = getLabel(tooltipSet);
        return { fieldRow: fieldRow, fieldInput: fieldInput };
    }

    /**
     * チェックボックス1つの行を作る
     * @param {Group} parentGroup - 追加先のグループ
     * @param {object} labelSet - チェックボックスのラベル定義
     * @param {object} tooltipSet - ツールチップ定義
     * @returns {{checkboxRow: Group, checkbox: Checkbox}} 作成した行とチェックボックス
     */
    function addCheckboxRow(parentGroup, labelSet, tooltipSet) {
        var checkboxRow = addRow(parentGroup);
        var rowCheckbox = checkboxRow.add("checkbox", undefined, getLabel(labelSet));
        rowCheckbox.helpTip = getLabel(tooltipSet);
        return { checkboxRow: checkboxRow, checkbox: rowCheckbox };
    }

    var incrementDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
    incrementDialog.orientation = "column";
    incrementDialog.alignChildren = "fill";

    var savedDialogPosition = readDialogPosition();
    if (savedDialogPosition) incrementDialog.location = savedDialogPosition;

    var settingsGroup = incrementDialog.add("group");
    settingsGroup.orientation = "column";
    settingsGroup.alignChildren = "left";

    /* 複製数 / Copies */
    var countInput = addFieldRow(settingsGroup, LABELS.fieldLabel.copyCount, String(DEFAULT_COPY_COUNT), NUMBER_FIELD_CHARS, LABELS.tooltip.copyCount).fieldInput;

    /* 増分 / Step */
    var stepInput = addFieldRow(settingsGroup, LABELS.fieldLabel.stepValue, String(DEFAULT_STEP), NUMBER_FIELD_CHARS, LABELS.tooltip.stepValue).fieldInput;

    /* アキ（文字サイズに足す量） / Gap added on top of the font size */
    var defaultGapPt = sourceFontSizePt * (DEFAULT_PITCH_RATIO - 1);
    if (defaultGapPt < 0) defaultGapPt = 0;
    var gapRow = addFieldRow(settingsGroup, LABELS.fieldLabel.gap, ptToUnitValue(defaultGapPt, textUnitInfo).toFixed(1), NUMBER_FIELD_CHARS, LABELS.tooltip.gap);
    var gapInput = gapRow.fieldInput;
    gapRow.fieldRow.add("statictext", undefined, textUnitInfo.label);

    /* 増分対象（候補が複数あるときだけ作る） / Increment target (only when there is a choice) */
    var targetRadios = [];
    if (targetRadioLabels.length > 1) {
        var targetRow = addRow(settingsGroup);
        addRowLabel(targetRow, LABELS.fieldLabel.incrementTarget);

        var targetRadioGroup = targetRow.add("group");
        targetRadioGroup.orientation = "row";
        targetRadioGroup.alignChildren = ["left", "center"];

        for (var i = 0; i < targetRadioLabels.length; i++) {
            var targetRadio = targetRadioGroup.add("radiobutton", undefined, targetRadioLabels[i]);
            targetRadio.value = (targetTokenIndices[i] === targetTokenIndex);
            targetRadio.helpTip = getLabel(LABELS.tooltip.incrementTarget);
            targetRadios.push(targetRadio);
        }
    }

    /* 開始値 / Start value */
    var startRow = addCheckboxRow(settingsGroup, LABELS.checkbox.startOverride, LABELS.tooltip.startOverride);
    var startOverrideCheckbox = startRow.checkbox;

    var startValueInput = startRow.checkboxRow.add("edittext", undefined, String(sourceTokens[targetTokenIndex]));
    startValueInput.characters = START_FIELD_CHARS;
    startValueInput.helpTip = getLabel(LABELS.tooltip.startValue);
    startValueInput.enabled = false;

    /* ゼロ埋め / Zero padding */
    var zeroPadCheckbox = addCheckboxRow(settingsGroup, LABELS.checkbox.zeroPad, LABELS.tooltip.zeroPad).checkbox;
    zeroPadCheckbox.value = DEFAULT_ZERO_PAD;

    /* 確定時にテキストを結合 / Merge text on OK */
    var mergeOnOKCheckbox = addCheckboxRow(settingsGroup, LABELS.checkbox.mergeOnOK, LABELS.tooltip.mergeOnOK).checkbox;
    mergeOnOKCheckbox.value = DEFAULT_MERGE_ON_OK;

    /* ボタンエリア / Buttons */
    var btnRowGroup = incrementDialog.add("group");
    btnRowGroup.orientation = "row";
    btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
    btnRowGroup.alignment = ["fill", "bottom"];

    var spacer = btnRowGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 0;

    var btnRightGroup = btnRowGroup.add("group");
    btnRightGroup.alignChildren = ["right", "center"];
    var btnCancel = btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
    var btnOK = btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

    // =========================================
    // イベント / Event handlers
    // =========================================

    /**
     * 増分対象のラジオが切り替わったときの処理
     * @param {number} radioIndex - 選択されたラジオの位置
     * @returns {void}
     */
    function onTargetChanged(radioIndex) {
        targetTokenIndex = targetTokenIndices[radioIndex];
        targetDigitLength = String(sourceTokens[targetTokenIndex]).length;
        startValueInput.text = String(sourceTokens[targetTokenIndex]);
        startValueInput.enabled = startOverrideCheckbox.value;
        updatePreview();
    }

    for (var j = 0; j < targetRadios.length; j++) {
        (function (radioIndex) {
            targetRadios[radioIndex].onClick = function () { onTargetChanged(radioIndex); };
        })(j);
    }

    startOverrideCheckbox.onClick = function () {
        startValueInput.enabled = startOverrideCheckbox.value;
        updatePreview();
    };
    zeroPadCheckbox.onClick = updatePreview;

    countInput.onChanging = updatePreview;
    gapInput.onChanging = updatePreview;
    stepInput.onChanging = updatePreview;
    startValueInput.onChanging = updatePreview;

    changeValueByArrowKey(countInput, false, updatePreview, true);
    changeValueByArrowKey(stepInput, true, updatePreview, true);
    changeValueByArrowKey(gapInput, false, updatePreview, false);
    changeValueByArrowKey(startValueInput, false, updatePreview, true);

    btnCancel.onClick = function () {
        incrementDialog.close();
    };

    btnOK.onClick = function () {
        closedWithOK = true;
        applySettingsToSourceText();

        /* プレビューがそのまま結果になる。空のときだけ作り直す / the preview is the result */
        if (previewItems.length === 0) previewItems = createIncrementedCopies();
        if (mergeOnOKCheckbox.value) mergeCopiesIntoSource();

        /* 複製は確定したので、プレビューの管理対象から外す / keep the copies in the document */
        previewItems = [];
        incrementDialog.close();
    };

    /* タイトルバーの×で閉じたときも、プレビューを消して元のテキストに戻す / Also covers the close box */
    incrementDialog.onClose = function () {
        saveDialogPosition(incrementDialog);
        if (closedWithOK) return;
        restoreSourceText();
    };

    updatePreview();
    incrementDialog.show();

})();
