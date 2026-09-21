#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した数字・英字・漢数字のテキストを、指定した並び順でソートして連番を振り直します。
書式は［基準となる値］のラジオで選び（123／ABC／abc／一二三／I II III／壱弐参）、ゼロ埋め、接頭辞・接尾辞の追加、番号順への重ね順の並べ替えにも対応します。

詳細は README を参照してください。

### Overview

Sorts the selected text — digits, letters, or Japanese numerals — in a chosen order and renumbers it as a sequence.
Radios pick the format (123 / ABC / abc / 一二三 / I II III / 壱弐参), with optional zero padding, a prefix or suffix, and restacking to match the new numbers.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartRenumber";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-12-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-21";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartRenumber.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartRenumber.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nf3b6601cd165"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 同じ行・同じ列と見なす座標の許容差（pt）。Z方向・N方向の並べ替えで使う
       Tolerance in points for treating items as one row or column (Z and N patterns) */
    var ROW_COLUMN_TOLERANCE_PT = 10;

    /* ［重ね順調整］をONで開くか / Whether Adjust Stacking starts checked */
    var REORDER_STACK_BY_DEFAULT = true;

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ダイアログの透明度と、入力欄の幅 / Dialog opacity and field widths */
    var DIALOG_LAYOUT = {
        opacity: 0.98,           /* ダイアログの透明度 / dialog opacity */
        startValueChars: 5,      /* 開始値の入力幅（文字数） / width of the start value field */
        groupTopMargin: 10,      /* パネル内で段を分けるときの上マージン / top margin inside a panel */
        affixChars: 8            /* 接頭辞・接尾辞の入力幅（文字数） / width of the affix fields */
    };

    // =========================================
    // 並び順 / Sort orders
    // =========================================

    /* 並び順の並べる順番。ラジオボタンの並び、LABELS.radio / LABELS.tooltip のキー、
       SORT_COMPARATORS のキーをこの名前で揃える
       Sort order sequence; the same keys are used for the radios, LABELS, and SORT_COMPARATORS */
    var SORT_MODES = ["currentValue", "horizontal", "vertical", "zPattern", "nPattern", "stackOrder"];

    /* 並び順ごとの比較関数。top は上にあるほど値が大きいので、降順で「上から下」になる
       Comparator per sort order; top grows upward, so descending means top-to-bottom */
    var SORT_COMPARATORS = {
        currentValue: function (entryA, entryB) {
            return entryA.sortValue - entryB.sortValue;
        },
        horizontal: function (entryA, entryB) {
            return entryA.left - entryB.left;
        },
        vertical: function (entryA, entryB) {
            return entryB.top - entryA.top;
        },
        zPattern: function (entryA, entryB) {
            /* 行が違うときだけ上下で比べ、同じ行なら左から右 / Compare rows first, then left to right */
            if (Math.abs(entryA.top - entryB.top) > ROW_COLUMN_TOLERANCE_PT) return entryB.top - entryA.top;
            return entryA.left - entryB.left;
        },
        nPattern: function (entryA, entryB) {
            /* 列が違うときだけ左右で比べ、同じ列なら上から下 / Compare columns first, then top to bottom */
            if (Math.abs(entryA.left - entryB.left) > ROW_COLUMN_TOLERANCE_PT) return entryA.left - entryB.left;
            return entryB.top - entryA.top;
        },
        stackOrder: function (entryA, entryB) {
            /* 前面が先。レイヤーをまたぐときはレイヤーの重ね順を先に見る
               Frontmost first; compare the layer order first when the layers differ */
            if (entryA.layerOrder !== entryB.layerOrder) return entryB.layerOrder - entryA.layerOrder;
            if (entryA.itemOrder !== entryB.itemOrder) return entryB.itemOrder - entryA.itemOrder;
            /* 重ね順が読めなかったときは選択順で安定させる / Fall back to the selection order */
            return entryA.selectionIndex - entryB.selectionIndex;
        }
    };

    // =========================================
    // 連番の書式 / Sequence format
    // =========================================

    /* ［基準となる値］の書式ラジオ。keyはLABELS.radioのキー、startValueは選んだときに入れる値
       Format radios; key matches LABELS.radio, startValue is dropped into the start value field */
    var FORMAT_MODES = [
        { key: "digit", startValue: "1" },
        { key: "upperCase", startValue: "A" },
        { key: "lowerCase", startValue: "a" },
        { key: "kanji", startValue: "一" },
        { key: "roman", startValue: "I" },
        { key: "daiji", startValue: "壱" }
    ];

    /* 英字だけのテキスト（A、Z、AA など） / Text made of letters only */
    var LETTERS_ONLY_PATTERN = /^[A-Za-z]+$/;

    /* 漢数字だけのテキスト（一、十二、弐、参拾 など）。通常と大字のどちらも受ける
       Text made of Japanese numerals only, plain or formal */
    var KANJI_ONLY_PATTERN = /^[〇零一二三四五六七八九十拾百千万萬壱弐参]+$/;

    /* 大字を含むか。含めば書き出しも大字にする / Formal numerals; their presence picks the formal output */
    var DAIJI_PATTERN = /[零壱弐参拾萬]/;

    /* 書き出しに使う漢数字。大字は一二三十を壱弐参拾に置き換える（法令で定められている4文字）
       Numerals used for output; the formal style replaces 一二三十 with 壱弐参拾 */
    var KANJI_STYLES = {
        kanji: {
            digits: ["〇", "一", "二", "三", "四", "五", "六", "七", "八", "九"],
            units: [
                { value: 1000, character: "千" },
                { value: 100, character: "百" },
                { value: 10, character: "十" }
            ]
        },
        daiji: {
            digits: ["零", "壱", "弐", "参", "四", "五", "六", "七", "八", "九"],
            units: [
                { value: 1000, character: "千" },
                { value: 100, character: "百" },
                { value: 10, character: "拾" }
            ]
        }
    };

    /* 読み取り用。通常・大字のどちらの表記も受ける / Reading maps covering both styles */
    var KANJI_VALUES = {
        "〇": 0, "零": 0, "一": 1, "壱": 1, "二": 2, "弐": 2, "三": 3, "参": 3,
        "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9
    };
    var KANJI_UNIT_VALUES = { "十": 10, "拾": 10, "百": 100, "千": 1000 };

    /* ローマ数字（I、IV、XII など）。入力は小文字でも受け付け、書き出しは大文字に揃える
       Roman numerals such as I, IV and XII; typed in either case, always written in capitals */
    var ROMAN_PATTERN = /^m*(cm|cd|d?c{0,3})(xc|xl|l?x{0,3})(ix|iv|v?i{0,3})$/i;

    /* ローマ数字の表記。大きい順に並べる / Roman numeral spellings, largest first */
    var ROMAN_UNITS = [
        { value: 1000, characters: "M" }, { value: 900, characters: "CM" },
        { value: 500, characters: "D" }, { value: 400, characters: "CD" },
        { value: 100, characters: "C" }, { value: 90, characters: "XC" },
        { value: 50, characters: "L" }, { value: 40, characters: "XL" },
        { value: 10, characters: "X" }, { value: 9, characters: "IX" },
        { value: 5, characters: "V" }, { value: 4, characters: "IV" },
        { value: 1, characters: "I" }
    ];

    /**
     * 前後の空白を落とします。
     *
     * @param {string} text - 対象の文字列。
     * @returns {string} 前後の空白を除いた文字列。
     */
    function trimText(text) {
        return String(text).replace(/^\s+|\s+$/g, "");
    }

    /**
     * 英字を1から始まる位置に換算します（A=1、Z=26、AA=27）。
     *
     * @param {string} letters - 英字だけの文字列。
     * @returns {number} 文字順の位置。
     */
    function lettersToIndex(letters) {
        var upperCaseLetters = letters.toUpperCase();
        var index = 0;
        for (var i = 0; i < upperCaseLetters.length; i++) {
            index = index * 26 + (upperCaseLetters.charCodeAt(i) - 64);
        }
        return index;
    }

    /**
     * 位置を英字に戻します（1=A、26=Z、27=AA）。
     *
     * @param {number} index - 1から始まる文字順の位置。
     * @returns {string} 大文字の英字。1以下は "A"。
     */
    function indexToLetters(index) {
        var letters = "";
        var remaining = Math.floor(index);
        while (remaining > 0) {
            var remainder = (remaining - 1) % 26;
            letters = String.fromCharCode(65 + remainder) + letters;
            remaining = Math.floor((remaining - 1) / 26);
        }
        return letters || "A";
    }

    /**
     * 漢数字を数に換算します（十二=12、二百三=203、一万=10000）。大字（壱弐参拾萬）も読めます。
     *
     * @param {string} kanjiText - 漢数字だけの文字列。
     * @returns {number} 換算した数。解釈できないときは null。
     */
    function kanjiToNumber(kanjiText) {
        var total = 0;      /* 万より上の確定分 / settled part above 10000 */
        var section = 0;    /* 千・百・十の合計 / sum of the thousand, hundred and ten places */
        var current = 0;    /* 直前の1桁 / the digit just read */
        var hasDigit = false;

        for (var i = 0; i < kanjiText.length; i++) {
            var character = kanjiText.charAt(i);

            if (KANJI_VALUES[character] !== undefined) {
                current = KANJI_VALUES[character];
                hasDigit = true;
                continue;
            }

            if (character === "万" || character === "萬") {
                total += (section + current) * 10000;
                section = 0;
                current = 0;
                continue;
            }

            var unitValue = KANJI_UNIT_VALUES[character];
            if (unitValue === undefined) return null;

            /* 「十」のように数字を伴わない位は1と見なす / A bare unit such as 十 counts as one */
            section += (current || 1) * unitValue;
            current = 0;
            hasDigit = true;
        }

        return hasDigit ? (total + section + current) : null;
    }

    /**
     * 数を漢数字に直します（12=十二、203=二百三、10000=一万）。大字では 12=拾弐 になります。
     *
     * @param {number} value - 0以上の整数。
     * @param {string} styleKey - "kanji"（通常）または "daiji"（大字）。
     * @returns {string} 漢数字。0以下は "〇"（大字は "零"）。
     */
    function numberToKanji(value, styleKey) {
        var style = KANJI_STYLES[styleKey] || KANJI_STYLES.kanji;
        var remaining = Math.floor(value);
        if (remaining <= 0) return style.digits[0];

        var kanji = "";

        /* 万の位から先に切り出す / Take the ten-thousands apart first */
        if (remaining >= 10000) {
            kanji += numberToKanji(Math.floor(remaining / 10000), styleKey) + "万";
            remaining = remaining % 10000;
            if (remaining === 0) return kanji;
        }

        for (var i = 0; i < style.units.length; i++) {
            var digit = Math.floor(remaining / style.units[i].value);
            if (digit > 0) {
                /* 十・百・千の「一」は書かない / The leading one is left out for 十, 百 and 千 */
                if (digit > 1) kanji += style.digits[digit];
                kanji += style.units[i].character;
                remaining -= digit * style.units[i].value;
            }
        }

        return kanji + ((remaining > 0) ? style.digits[remaining] : "");
    }

    /**
     * 並べ替えに使う値を返します。数字はその値、英字は文字順の位置、漢数字は換算した数です。
     *
     * @param {string} textContents - テキストの中身。
     * @returns {number} 並べ替え用の値。対象外のときは null。
     */
    function getSortValue(textContents) {
        var trimmedText = trimText(textContents);
        if (trimmedText === "") return null;
        if (LETTERS_ONLY_PATTERN.test(trimmedText)) return lettersToIndex(trimmedText);
        if (KANJI_ONLY_PATTERN.test(trimmedText)) return kanjiToNumber(trimmedText);
        if (!isNaN(parseFloat(trimmedText)) && isFinite(trimmedText)) return parseFloat(trimmedText);
        return null;
    }

    /**
     * ［開始値］を解釈します。数字（0や負の値も可）、英字、漢数字を受け付けます。
     *
     * @param {string} startValueText - ［開始値］の入力。
     * @param {string} formatModeKey - 選ばれている書式ラジオのキー（紛らわしい表記の判断に使う）。
     * @returns {object} format（"number" / "letter" / "kanji" / "daiji" / "roman"）と値を持つオブジェクト。
     *                   数字のときは入力した整数部の桁数を digits に入れます。解釈できないときは null。
     */
    function parseStartValue(startValueText, formatModeKey) {
        var trimmedText = trimText(startValueText);
        if (trimmedText === "") return null;

        /* I や X は英字とも読めるため、ローマ数字かどうかはラジオの選択で決める
           Letters like I and X are ambiguous, so the radio decides whether it is Roman */
        if (formatModeKey === "roman" && ROMAN_PATTERN.test(trimmedText)) {
            return { format: "roman", number: romanToNumber(trimmedText) };
        }

        if (LETTERS_ONLY_PATTERN.test(trimmedText)) {
            return {
                format: "letter",
                index: lettersToIndex(trimmedText),
                isUpperCase: (trimmedText === trimmedText.toUpperCase())
            };
        }

        if (KANJI_ONLY_PATTERN.test(trimmedText)) {
            var kanjiNumber = kanjiToNumber(trimmedText);
            if (kanjiNumber === null) return null;
            /* 四〜九は通常と大字で同じ字なので、判別できないときはラジオの選択に従う
               四 to 九 are shared by both styles, so the radio decides when the text cannot tell */
            var isDaiji = DAIJI_PATTERN.test(trimmedText) ||
                (formatModeKey === "daiji" && !/[〇一二三十]/.test(trimmedText));
            return { format: isDaiji ? "daiji" : "kanji", number: kanjiNumber };
        }

        var startNumber = parseFloat(trimmedText);
        if (isNaN(startNumber) || !isFinite(trimmedText)) return null;

        /* 入力した整数部の桁数を控える。"01" なら 2 で、以降も2桁で書き出す
           Remember how many integer digits were typed: "01" means two, and stays two */
        var integerDigits = /^[-+]?(\d+)/.exec(trimmedText);
        return {
            format: "number",
            number: startNumber,
            digits: integerDigits ? integerDigits[1].length : 1
        };
    }

    /**
     * ローマ数字を数に換算します（IV=4、XII=12）。
     *
     * @param {string} romanText - ローマ数字の文字列。
     * @returns {number} 換算した数。
     */
    function romanToNumber(romanText) {
        var upperCaseText = romanText.toUpperCase();
        var total = 0;
        var position = 0;

        while (position < upperCaseText.length) {
            for (var i = 0; i < ROMAN_UNITS.length; i++) {
                var characters = ROMAN_UNITS[i].characters;
                if (upperCaseText.substr(position, characters.length) === characters) {
                    total += ROMAN_UNITS[i].value;
                    position += characters.length;
                    break;
                }
            }
        }

        return total;
    }

    /**
     * 数をローマ数字に直します（4=IV、12=XII）。
     *
     * @param {number} value - 1以上の整数。
     * @returns {string} 大文字のローマ数字。1未満は "I"。
     */
    function numberToRoman(value) {
        var remaining = Math.floor(value);
        if (remaining < 1) return "I";

        var roman = "";
        for (var i = 0; i < ROMAN_UNITS.length; i++) {
            while (remaining >= ROMAN_UNITS[i].value) {
                roman += ROMAN_UNITS[i].characters;
                remaining -= ROMAN_UNITS[i].value;
            }
        }
        return roman;
    }

    /**
     * ［開始値］に対応する書式ラジオのキーを返します。
     *
     * @param {object} startValue - parseStartValue() の戻り値。
     * @returns {string} FORMAT_MODES のキー。判定できないときは null。
     */
    function getFormatModeKey(startValue) {
        if (!startValue) return null;
        if (startValue.format === "roman") return "roman";
        if (startValue.format === "daiji") return "daiji";
        if (startValue.format === "kanji") return "kanji";
        if (startValue.format === "letter") return startValue.isUpperCase ? "upperCase" : "lowerCase";
        return "digit";
    }

    /**
     * ［開始値］から数えて offset 番目の値を文字列で返します。
     *
     * @param {object} startValue - parseStartValue() の戻り値。
     * @param {number} offset - ［開始値］からの位置（0が最初）。
     * @param {number} zeroPadDigits - ゼロ埋めの桁数。0で埋めません。
     * @returns {string} 書き込む文字列。
     */
    function formatSequenceValue(startValue, offset, zeroPadDigits) {
        if (startValue.format === "letter") {
            var letters = indexToLetters(startValue.index + offset);
            return startValue.isUpperCase ? letters : letters.toLowerCase();
        }

        if (startValue.format === "kanji" || startValue.format === "daiji") {
            return numberToKanji(startValue.number + offset, startValue.format);
        }

        if (startValue.format === "roman") {
            return numberToRoman(startValue.number + offset);
        }

        var assignedNumber = startValue.number + offset;
        return (zeroPadDigits > 1) ? padWithZeros(assignedNumber, zeroPadDigits) : String(assignedNumber);
    }

    /**
     * 書き出す整数部の桁数を返します。入力した桁数（"01" なら2桁）は常に保ち、
     * ［ゼロ埋め］がONのときは最後の番号の桁数まで広げます。
     *
     * @param {object} startValue - parseStartValue() の戻り値。
     * @param {number} count - 振り直す個数。
     * @param {boolean} isZeroPadded - ［ゼロ埋め］がONか。
     * @returns {number} ゼロ埋めする桁数。埋めないときは0。
     */
    function getSequenceDigits(startValue, count, isZeroPadded) {
        if (startValue.format !== "number") return 0;
        if (!isZeroPadded) return startValue.digits;
        return Math.max(startValue.digits, getZeroPadDigits(startValue.number, count));
    }

    /**
     * ［ゼロ埋め］をONにすると桁数が変わるかを返します。ディム判定に使います。
     *
     * @param {object} startValue - parseStartValue() の戻り値。
     * @param {number} count - 振り直す個数。
     * @returns {boolean} 桁数が変わるなら true。
     */
    function willZeroPadApply(startValue, count) {
        if (!startValue || startValue.format !== "number") return false;
        return getZeroPadDigits(startValue.number, count) > startValue.digits;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境のロケールからUI言語を判定します。
     *
     * @returns {string} "ja" または "en"。
     */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* UI文言の定義 / UI string definitions */
    var LABELS = {
        dialog: {
            title: { ja: "連番振り直し", en: "Renumber Sequence" }
        },
        panel: {
            baseValue: { ja: "基準となる値", en: "Base Value" },
            sortOrder: { ja: "並び順", en: "Sort Order" },
            options: { ja: "オプション", en: "Options" },
            textAdd: { ja: "テキスト追加", en: "Text Addition" }
        },
        fieldLabel: {
            startValue: { ja: "開始値", en: "Start Value" },
            prefix: { ja: "接頭辞", en: "Prefix" },
            suffix: { ja: "接尾辞", en: "Suffix" }
        },
        radio: {
            digit: { ja: "123", en: "123" },
            upperCase: { ja: "ABC", en: "ABC" },
            lowerCase: { ja: "abc", en: "abc" },
            kanji: { ja: "一二三", en: "一二三" },
            roman: { ja: "I II III", en: "I II III" },
            daiji: { ja: "壱弐参", en: "壱弐参" },
            currentValue: { ja: "現在の値順", en: "Current Value Order" },
            horizontal: { ja: "水平方向（左から右）", en: "Horizontal (Left to Right)" },
            vertical: { ja: "垂直方向（上から下）", en: "Vertical (Top to Bottom)" },
            zPattern: { ja: "Z方向（左→右、上→下）", en: "Z-Pattern (Left-to-Right, Row-major)" },
            nPattern: { ja: "N方向（上→下、左→右）", en: "N-Pattern (Top-to-Bottom, Column-major)" },
            stackOrder: { ja: "重ね順（前面から）", en: "Stacking Order (Front to Back)" }
        },
        checkbox: {
            reverse: { ja: "逆順", en: "Reverse" },
            zeroPad: { ja: "ゼロ埋め", en: "Zero Padding" },
            reorderStack: { ja: "重ね順調整（OK時）", en: "Adjust Stacking (on OK)" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document open." },
            noSelection: { ja: "テキストオブジェクトを選択してください。", en: "Please select text objects." },
            noTargetText: {
                ja: "数字・英字・漢数字だけのテキストオブジェクトが見つかりませんでした。",
                en: "No text objects containing only a number, letters, or Japanese numerals were found."
            }
        },
        tooltip: {
            formatMode: {
                ja: "書き出す書式を選びます。選ぶと［開始値］に 1／A／a／一／I／壱 が入ります。",
                en: "Picks the format; choosing one drops 1, A, a, 一, I, or 壱 into the start value."
            },
            startValue: {
                ja: "振り直しの最初の値です。数字（1、0、-3）、英字（A、a）、漢数字（一）、ローマ数字（I）、大字（壱）を指定できます。",
                en: "First value: a number (1, 0, -3), a letter (A, a), a Japanese numeral (一 or 壱), or a Roman numeral (I)."
            },
            startValueKeys: {
                ja: "↑↓で±1、Shift＋↑↓で±10、Option＋↑↓で±0.1（英字は↑↓で1つずつ）",
                en: "Up/Down: ±1, Shift+Up/Down: ±10, Option+Up/Down: ±0.1 (letters step one at a time)"
            },
            currentValue: {
                ja: "いま入っている値の小さい順に振り直します（英字はアルファベット順）。",
                en: "Renumbers in ascending order of the current values; letters go in alphabetical order."
            },
            horizontal: {
                ja: "左にあるものから順に振り直します。",
                en: "Renumbers from the leftmost object to the right."
            },
            vertical: { ja: "上にあるものから順に振り直します。", en: "Renumbers from the topmost object down." },
            zPattern: {
                ja: "行ごとに左から右へ、行を上から下へ進みます。",
                en: "Moves left to right within a row, then down to the next row."
            },
            nPattern: {
                ja: "列ごとに上から下へ、列を左から右へ進みます。",
                en: "Moves top to bottom within a column, then right to the next column."
            },
            stackOrder: {
                ja: "重ね順の前面にあるものから順に振り直します。",
                en: "Renumbers from the frontmost object backward."
            },
            reverse: { ja: "並び順を逆さにして振り直します。", en: "Reverses the order before renumbering." },
            zeroPad: {
                ja: "いちばん桁数の多い番号に合わせて、頭に0を足します。入力した桁数（01なら2桁）は、このチェックに関わらず保ちます。",
                en: "Pads with leading zeros to match the widest number. The typed width (two digits for 01) is kept either way."
            },
            reorderStack: {
                ja: "振り直した番号の順に重ね順を並べ替えます（番号の小さいものが前面）。［OK］のときだけ実行します。",
                en: "Reorders the stacking to follow the new numbers, lowest number frontmost. Applied on OK only."
            },
            prefix: { ja: "番号の前に付ける文字列です。", en: "Text placed before the number." },
            suffix: { ja: "番号の後ろに付ける文字列です。", en: "Text placed after the number." }
        }
    };

    /**
     * 言語に合わせた文言を返します。
     *
     * @param {object} labelSet - ja / en を持つ文言オブジェクト。
     * @returns {string} 現在の言語の文言。
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || "";
    }

    /**
     * 項目名にコロンを付けて返します（日本語は全角、英語は半角）。
     *
     * @param {object} labelSet - ja / en を持つ文言オブジェクト。
     * @returns {string} コロン付きの項目名。
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + ((uiLang === "ja") ? "：" : ": ");
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    main();

    /**
     * ドキュメントと選択を確認し、振り直しのダイアログを開きます。
     *
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        if (!selectedItems || selectedItems.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var renumberTargets = collectSequenceTargets(selectedItems);
        if (renumberTargets.length === 0) {
            alert(getLabel(LABELS.alert.noTargetText));
            return;
        }

        showRenumberDialog(renumberTargets);
    }

    /**
     * 選択の中から、数字・英字・漢数字だけが入ったテキストオブジェクトを集めます。
     *
     * 並べ替えに使う座標と重ね順は、プレビューで中身が変わる前に控えます。
     *
     * @param {Array} selectedItems - ドキュメントの選択。
     * @returns {Array<object>} 振り直し対象（frame / originalContents / text / sortValue / left / top /
     *                   layerOrder / itemOrder / selectionIndex）。
     */
    function collectSequenceTargets(selectedItems) {
        var renumberTargets = [];

        for (var i = 0; i < selectedItems.length; i++) {
            var item = selectedItems[i];
            if (item.typename !== "TextFrame") continue;

            /* 中身が数字・英字・漢数字だけのものに限る（"3行目" などは対象外）
               Only text that is nothing but digits, letters, or Japanese numerals */
            var sortValue = getSortValue(item.contents);
            if (sortValue === null) continue;

            var stackOrderKey = getStackOrderKey(item);
            renumberTargets.push({
                frame: item,
                originalContents: item.contents,
                text: trimText(item.contents),
                sortValue: sortValue,
                left: item.left,
                top: item.top,
                layerOrder: stackOrderKey.layerOrder,
                itemOrder: stackOrderKey.itemOrder,
                selectionIndex: i
            });
        }

        return renumberTargets;
    }

    /**
     * 重ね順の比較に使うキーを返します。値が大きいほど前面です。
     *
     * @param {PageItem} item - 対象のオブジェクト。
     * @returns {{layerOrder: number, itemOrder: number}} レイヤーとレイヤー内の重ね順。
     */
    function getStackOrderKey(item) {
        /* zOrderPosition を持たないアイテムもあるため、まとめて保護する
           Some items expose no zOrderPosition, so guard the whole lookup */
        try {
            return { layerOrder: item.layer.zOrderPosition, itemOrder: item.zOrderPosition };
        } catch (e) {
            return { layerOrder: 0, itemOrder: 0 };
        }
    }

    /**
     * 連番振り直しのダイアログを表示します。
     *
     * @param {Array<object>} renumberTargets - 振り直し対象。
     * @returns {void}
     */
    function showRenumberDialog(renumberTargets) {
        /* ［OK］で閉じたかどうか。×やキャンセルのときだけ元に戻す
           Whether OK was used; only a cancel or a close restores the originals */
        var isConfirmed = false;

        var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.spacing = 10;
        dialog.margins = 20;
        dialog.opacity = DIALOG_LAYOUT.opacity;

        /* 上段：2カラム / Upper area: two columns */
        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = 20;

        /* 左カラム：基準となる値とオプション / Left column: base value and options */
        var leftColumn = columnsGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];
        leftColumn.spacing = 15;

        var baseValueControls = addStartValuePanel(leftColumn, getInitialStartValue(renumberTargets));
        var startValueInput = baseValueControls.startValueInput;
        var formatRadios = baseValueControls.formatRadios;
        var renumberOptions = addOptionsPanel(leftColumn);

        /* 右カラム：並び順とテキスト追加 / Right column: sort order and affixes */
        var rightColumn = columnsGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "top"];
        rightColumn.spacing = 15;

        var sortOrderControls = addSortOrderPanel(rightColumn);
        var sortOrderRadios = sortOrderControls.radios;
        var affixInputs = addAffixPanel(rightColumn);

        /* 下段：ボタンエリア / Lower area: buttons */
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["fill", "top"];
        btnRowGroup.alignChildren = ["fill", "center"];
        btnRowGroup.margins = [0, 5, 0, 0];
        btnRowGroup.spacing = 0;

        /* 横に伸びる空白でボタンを右へ寄せる / Spacer pushes the buttons to the right */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;
        spacer.maximumSize.height = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.orientation = "row";
        btnRightGroup.alignment = ["right", "center"];
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = 10;

        var btnCancel = btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        /* 入力値をまとめて読み取る / Read every input at once */
        function getRenumberSettings() {
            return {
                startValueText: startValueInput.text,
                formatModeKey: getSelectedFormatMode(formatRadios),
                prefix: affixInputs.prefix.text,
                suffix: affixInputs.suffix.text,
                sortMode: getSelectedSortMode(sortOrderRadios),
                isReversed: renumberOptions.reverse.value,
                isZeroPadded: renumberOptions.zeroPad.value,
                isStackReordered: sortOrderControls.reorderStack.value
            };
        }

        /* プレビュー更新 / Refresh the preview */
        function updatePreview() {
            var startValue = parseStartValue(startValueInput.text, getSelectedFormatMode(formatRadios));

            /* 入力に合わせて書式のラジオを合わせる / Keep the format radios in step with the input */
            var formatModeKey = getFormatModeKey(startValue);
            if (formatModeKey) formatRadios[formatModeKey].value = true;

            /* ゼロ埋めが効かない書式・桁数のときはディム / Dim zero padding when it would change nothing */
            renumberOptions.zeroPad.enabled = willZeroPadApply(startValue, renumberTargets.length);

            /* ［開始値］を読めないあいだは元の中身に戻しておく
               While the start value cannot be read, put the original contents back */
            if (!applyRenumber(renumberTargets, getRenumberSettings())) {
                restoreOriginalContents(renumberTargets);
            }
            app.redraw();
        }

        btnCancel.onClick = function () {
            dialog.close(0);
        };

        btnOK.onClick = function () {
            /* 元に戻してから一度だけ適用する。最後の書き込みがまとまるので、Undoで元の状態まで戻せる
               Restore first, then apply once, so undo lands back on the original text */
            restoreOriginalContents(renumberTargets);

            var settings = getRenumberSettings();
            var orderedTargets = applyRenumber(renumberTargets, settings);
            /* 重ね順の並べ替えはプレビューでは行わない（控えた重ね順が狂うため）
               Restacking runs here only: doing it in the preview would stale the cached order */
            if (orderedTargets && settings.isStackReordered) reorderStackToMatch(orderedTargets);

            isConfirmed = true;
            app.redraw();
            dialog.close(1);
        };

        /* キャンセル・×で閉じたときは元の中身に戻す / A cancel or a close puts the originals back */
        dialog.onClose = function () {
            if (!isConfirmed) {
                restoreOriginalContents(renumberTargets);
                app.redraw();
            }
            return true;
        };

        /* ラジオを選んだら、その書式のいちばん若い値を［開始値］に入れる
           Choosing a format drops its lowest value into the start value field */
        function makeFormatClickHandler(formatStartValue) {
            return function () {
                startValueInput.text = formatStartValue;
                updatePreview();
            };
        }
        for (var f = 0; f < FORMAT_MODES.length; f++) {
            formatRadios[FORMAT_MODES[f].key].onClick = makeFormatClickHandler(FORMAT_MODES[f].startValue);
        }

        startValueInput.onChanging = updatePreview;
        affixInputs.prefix.onChanging = updatePreview;
        affixInputs.suffix.onChanging = updatePreview;
        renumberOptions.reverse.onClick = updatePreview;
        renumberOptions.zeroPad.onClick = updatePreview;
        for (var i = 0; i < SORT_MODES.length; i++) {
            sortOrderRadios[SORT_MODES[i]].onClick = updatePreview;
        }

        updatePreview();
        dialog.show();
    }

    /**
     * ［基準となる値］パネル（書式のラジオと開始値）を作ります。
     *
     * @param {Group} parentGroup - 追加先のグループ。
     * @param {string} initialStartValue - 初期値。
     * @returns {{startValueInput: EditText, formatRadios: object}} 開始値の入力欄と書式のラジオ。
     */
    function addStartValuePanel(parentGroup, initialStartValue) {
        var baseValuePanel = parentGroup.add("panel", undefined, getLabel(LABELS.panel.baseValue));
        baseValuePanel.orientation = "column";
        baseValuePanel.alignChildren = ["left", "top"];
        baseValuePanel.margins = [15, 20, 15, 10];
        baseValuePanel.spacing = 4;

        /* 書式のラジオ（縦並び）。同じ親に入れておくと排他になる
           Format radios in one column; keeping them in one parent keeps them exclusive */
        var formatColumn = baseValuePanel.add("group");
        formatColumn.orientation = "column";
        formatColumn.alignChildren = ["left", "top"];
        formatColumn.spacing = 4;

        var formatRadios = {};
        for (var i = 0; i < FORMAT_MODES.length; i++) {
            var formatRadio = formatColumn.add("radiobutton", undefined, getLabel(LABELS.radio[FORMAT_MODES[i].key]));
            formatRadio.helpTip = getLabel(LABELS.tooltip.formatMode);
            formatRadios[FORMAT_MODES[i].key] = formatRadio;
        }

        /* 項目名を上、入力欄を下に置く。ラジオとのアキは上マージンで取る
           The label sits above the field, with a top margin separating it from the radios */
        var startValueGroup = baseValuePanel.add("group");
        startValueGroup.orientation = "column";
        startValueGroup.alignChildren = ["left", "top"];
        startValueGroup.margins = [0, DIALOG_LAYOUT.groupTopMargin, 0, 0];
        startValueGroup.spacing = 4;

        startValueGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.startValue));

        var startValueInput = startValueGroup.add("edittext", undefined, initialStartValue);
        startValueInput.characters = DIALOG_LAYOUT.startValueChars;
        startValueInput.helpTip = getLabel(LABELS.tooltip.startValue) + "\n" + getLabel(LABELS.tooltip.startValueKeys);
        startValueInput.active = true;
        /* 数字は±1／±10／±0.1、英字と漢数字は1つずつ / Digits step by 1, 10 or 0.1; letters and kanji by one */
        changeValueByArrowKey(startValueInput, true);
        changeSequenceValueByArrowKey(startValueInput, formatRadios);

        return { startValueInput: startValueInput, formatRadios: formatRadios };
    }

    /**
     * ［並び順］パネル（並び順のラジオと重ね順調整）を作ります。
     *
     * @param {Group} parentGroup - 追加先のグループ。
     * @returns {{radios: object, reorderStack: Checkbox}} 並び順のラジオと重ね順調整のチェックボックス。
     */
    function addSortOrderPanel(parentGroup) {
        var sortOrderPanel = parentGroup.add("panel", undefined, getLabel(LABELS.panel.sortOrder));
        sortOrderPanel.orientation = "column";
        sortOrderPanel.alignChildren = ["left", "top"];
        sortOrderPanel.margins = [15, 20, 15, 10];
        sortOrderPanel.spacing = 4;

        var sortOrderRadios = {};
        for (var i = 0; i < SORT_MODES.length; i++) {
            var sortMode = SORT_MODES[i];
            var sortModeRadio = sortOrderPanel.add("radiobutton", undefined, getLabel(LABELS.radio[sortMode]));
            sortModeRadio.helpTip = getLabel(LABELS.tooltip[sortMode]);
            sortOrderRadios[sortMode] = sortModeRadio;
        }
        sortOrderRadios[SORT_MODES[0]].value = true;

        /* 並び順に合わせて重ね順も揃えるので、ラジオの下に置く
           Sits under the radios: it lines the stacking up with the chosen order */
        var reorderStackGroup = sortOrderPanel.add("group");
        reorderStackGroup.orientation = "column";
        reorderStackGroup.alignChildren = ["left", "top"];
        reorderStackGroup.margins = [0, DIALOG_LAYOUT.groupTopMargin, 0, 0];

        var reorderStackCheckbox = reorderStackGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.reorderStack));
        reorderStackCheckbox.helpTip = getLabel(LABELS.tooltip.reorderStack);
        reorderStackCheckbox.value = REORDER_STACK_BY_DEFAULT;

        return { radios: sortOrderRadios, reorderStack: reorderStackCheckbox };
    }

    /**
     * 選択されている書式のキーを返します。
     *
     * @param {object} formatRadios - addStartValuePanel() が返したラジオボタン。
     * @returns {string} FORMAT_MODES のいずれか。
     */
    function getSelectedFormatMode(formatRadios) {
        for (var i = 0; i < FORMAT_MODES.length; i++) {
            if (formatRadios[FORMAT_MODES[i].key].value) return FORMAT_MODES[i].key;
        }
        return FORMAT_MODES[0].key;
    }

    /**
     * 選択されている並び順のキーを返します。
     *
     * @param {object} sortOrderRadios - addSortOrderPanel() が返したラジオボタン。
     * @returns {string} SORT_MODES のいずれか。
     */
    function getSelectedSortMode(sortOrderRadios) {
        for (var i = 0; i < SORT_MODES.length; i++) {
            if (sortOrderRadios[SORT_MODES[i]].value) return SORT_MODES[i];
        }
        return SORT_MODES[0];
    }

    /**
     * ［オプション］パネル（逆順・ゼロ埋め）を作ります。
     *
     * @param {Group} parentGroup - 追加先のグループ。
     * @returns {{reverse: Checkbox, zeroPad: Checkbox}} 作成したチェックボックス。
     */
    function addOptionsPanel(parentGroup) {
        var optionsPanel = parentGroup.add("panel", undefined, getLabel(LABELS.panel.options));
        optionsPanel.orientation = "column";
        optionsPanel.alignChildren = ["left", "top"];
        optionsPanel.margins = [15, 20, 15, 10];
        optionsPanel.spacing = 4;

        var reverseCheckbox = optionsPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.reverse));
        reverseCheckbox.helpTip = getLabel(LABELS.tooltip.reverse);

        var zeroPadCheckbox = optionsPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.zeroPad));
        zeroPadCheckbox.helpTip = getLabel(LABELS.tooltip.zeroPad);

        return { reverse: reverseCheckbox, zeroPad: zeroPadCheckbox };
    }

    /**
     * ［テキスト追加］パネル（接頭辞・接尾辞）を作ります。
     *
     * @param {Group} parentGroup - 追加先のグループ。
     * @returns {{prefix: EditText, suffix: EditText}} 作成した入力欄。
     */
    function addAffixPanel(parentGroup) {
        var affixPanel = parentGroup.add("panel", undefined, getLabel(LABELS.panel.textAdd));
        affixPanel.orientation = "column";
        affixPanel.alignChildren = ["right", "top"];
        affixPanel.margins = [15, 20, 15, 15];
        affixPanel.spacing = 8;

        return {
            prefix: addAffixRow(affixPanel, LABELS.fieldLabel.prefix, LABELS.tooltip.prefix),
            suffix: addAffixRow(affixPanel, LABELS.fieldLabel.suffix, LABELS.tooltip.suffix)
        };
    }

    /**
     * 接頭辞・接尾辞の1行を作ります。
     *
     * @param {Panel} parentPanel - 追加先のパネル。
     * @param {object} fieldLabelSet - 項目名の文言オブジェクト。
     * @param {object} tooltipSet - tooltipの文言オブジェクト。
     * @returns {EditText} 作成した入力欄。
     */
    function addAffixRow(parentPanel, fieldLabelSet, tooltipSet) {
        var affixRow = parentPanel.add("group");
        affixRow.add("statictext", undefined, labelText(fieldLabelSet));

        var affixInput = affixRow.add("edittext", undefined, "");
        affixInput.characters = DIALOG_LAYOUT.affixChars;
        affixInput.helpTip = getLabel(tooltipSet);

        return affixInput;
    }

    /**
     * 並べ替えた順に連番を書き込みます。プレビューと確定で共用します。
     *
     * @param {Array<object>} renumberTargets - 振り直し対象。
     * @param {object} settings - getRenumberSettings() が返す設定。
     * @returns {Array<object>} 番号を振った順に並べた配列。［開始値］が不正なときは null。
     */
    function applyRenumber(renumberTargets, settings) {
        var startValue = parseStartValue(settings.startValueText, settings.formatModeKey);
        if (!startValue) return null;

        var orderedTargets = sortTargets(renumberTargets, settings.sortMode, settings.isReversed);
        var zeroPadDigits = getSequenceDigits(startValue, orderedTargets.length, settings.isZeroPadded);

        for (var i = 0; i < orderedTargets.length; i++) {
            var sequenceText = formatSequenceValue(startValue, i, zeroPadDigits);
            orderedTargets[i].frame.contents = settings.prefix + sequenceText + settings.suffix;
        }

        return orderedTargets;
    }

    /**
     * 振り直した番号の順に重ね順を並べ替えます（番号の小さいものが前面）。
     *
     * @param {Array<object>} orderedTargets - 番号を振った順に並んだ対象。
     * @returns {void}
     */
    function reorderStackToMatch(orderedTargets) {
        /* 後ろから順に最前面へ送ると、先頭＝いちばん若い番号が最前面になる
           Bring each one to the front starting from the last, so the lowest number ends up frontmost */
        for (var i = orderedTargets.length - 1; i >= 0; i--) {
            orderedTargets[i].frame.zOrder(ZOrderMethod.BRINGTOFRONT);
        }
    }

    /**
     * 指定した並び順に並べ替えた配列を返します（元の配列は変更しません）。
     *
     * @param {Array<object>} renumberTargets - 振り直し対象。
     * @param {string} sortMode - SORT_MODES のいずれか。
     * @param {boolean} isReversed - 並びを逆さにするか。
     * @returns {Array<object>} 並べ替えた配列。
     */
    function sortTargets(renumberTargets, sortMode, isReversed) {
        var orderedTargets = renumberTargets.slice();
        orderedTargets.sort(SORT_COMPARATORS[sortMode] || SORT_COMPARATORS.currentValue);
        if (isReversed) orderedTargets.reverse();
        return orderedTargets;
    }

    /**
     * ゼロ埋めの桁数を返します。最初と最後の番号のうち、桁数の多いほうに合わせます。
     *
     * @param {number} startNumber - 開始番号。
     * @param {number} count - 振り直す個数。
     * @returns {number} 整数部の桁数。
     */
    function getZeroPadDigits(startNumber, count) {
        var firstDigits = countIntegerDigits(startNumber);
        var lastDigits = countIntegerDigits(startNumber + count - 1);
        return Math.max(firstDigits, lastDigits);
    }

    /**
     * 数値の整数部の桁数を返します（符号は数えません）。
     *
     * @param {number} num - 対象の数値。
     * @returns {number} 桁数。
     */
    function countIntegerDigits(num) {
        return String(Math.floor(Math.abs(num))).length;
    }

    /**
     * 整数部を指定の桁数までゼロ埋めした文字列を返します。
     *
     * @param {number} num - 対象の数値。
     * @param {number} digits - 整数部の桁数。
     * @returns {string} ゼロ埋めした文字列。
     */
    function padWithZeros(num, digits) {
        var sign = (num < 0) ? "-" : "";
        var absoluteText = String(Math.abs(num));
        var dotIndex = absoluteText.indexOf(".");
        var integerText = (dotIndex === -1) ? absoluteText : absoluteText.substring(0, dotIndex);
        var decimalText = (dotIndex === -1) ? "" : absoluteText.substring(dotIndex);

        while (integerText.length < digits) {
            integerText = "0" + integerText;
        }
        return sign + integerText + decimalText;
    }

    /**
     * ［開始値］の初期値を返します。いちばん小さい値のテキストをそのまま使うので、
     * 選択が英字なら英字、漢数字なら漢数字で始まります。
     *
     * @param {Array<object>} renumberTargets - 振り直し対象。
     * @returns {string} 初期値のテキスト。
     */
    function getInitialStartValue(renumberTargets) {
        var smallestTarget = renumberTargets[0];
        for (var i = 1; i < renumberTargets.length; i++) {
            if (renumberTargets[i].sortValue < smallestTarget.sortValue) smallestTarget = renumberTargets[i];
        }
        return smallestTarget.text;
    }

    /**
     * 控えておいた元の中身をすべて書き戻します。
     *
     * app.undo() は「1回の addStep ＝ Undo 1段」が前提で、書き込みが起きなかったときに
     * 段数がずれてユーザーの操作まで取り消してしまうため、控えた文字列で戻します。
     *
     * @param {Array<object>} renumberTargets - 振り直し対象。
     * @returns {void}
     */
    function restoreOriginalContents(renumberTargets) {
        for (var i = 0; i < renumberTargets.length; i++) {
            renumberTargets[i].frame.contents = renumberTargets[i].originalContents;
        }
    }

    /**
     * 英字・漢数字・ローマ数字の入力欄で↑↓キーによる増減を有効にします（A→B、Z→AA、十→十一、III→IV）。
     *
     * 数字のときは何もしません（changeValueByArrowKey() が処理します）。
     *
     * @param {EditText} editText - 対象の入力欄。
     * @param {object} formatRadios - 書式のラジオ（紛らわしい表記の判断に使う）。
     * @returns {void}
     */
    function changeSequenceValueByArrowKey(editText, formatRadios) {
        editText.addEventListener("keydown", function (event) {
            if (!event || (event.keyName !== "Up" && event.keyName !== "Down")) return;

            var startValue = parseStartValue(editText.text, getSelectedFormatMode(formatRadios));
            if (!startValue || startValue.format === "number") return;

            event.preventDefault();
            editText.text = formatSequenceValue(startValue, (event.keyName === "Up") ? 1 : -1, 0);

            /* keydownでtextを書き換えるとonChangingが発火しないことがあるため明示的に呼ぶ
               Call onChanging by hand: rewriting text from keydown does not always fire it */
            if (typeof editText.onChanging === "function") editText.onChanging();
        });
    }

    /**
     * 数値入力欄で↑↓キーによる増減を有効にします。
     *
     * @param {EditText} editText - 対象の入力欄。
     * @param {boolean} allowNegative - マイナスの値を許すか。
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative) {
        editText.addEventListener("keydown", function (event) {
            if (!event || (event.keyName !== "Up" && event.keyName !== "Down")) return;

            /* 空欄のときは何もしない（Number("") は 0 になってしまう）
               Do nothing on an empty field: Number("") would come out as 0 */
            var trimmedText = trimText(editText.text);
            if (trimmedText === "") return;

            var value = Number(trimmedText);
            if (isNaN(value)) return;

            /* 先にキーの既定動作を止める（値を書き換えたあとでは間に合わない環境がある）
               Cancel the default first: some hosts apply it before we finish */
            event.preventDefault();

            var keyboard = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName === "Up");
            var isFineStep = !!keyboard.altKey;

            if (keyboard.shiftKey) {
                /* Shiftキー押下時は10の倍数にスナップ / Snap to the nearest ten */
                value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
            } else {
                var delta = isFineStep ? 0.1 : 1;
                value = isUp ? (value + delta) : (value - delta);
            }

            if (!allowNegative && value < 0) value = 0;

            /* optionキー押下時は小数第1位まで、それ以外は整数に丸める
               Round to one decimal with option held, otherwise to an integer */
            value = isFineStep ? (Math.round(value * 10) / 10) : Math.round(value);

            editText.text = String(value);

            /* keydownでtextを書き換えるとonChangingが発火しないことがあるため明示的に呼ぶ
               Call onChanging by hand: rewriting text from keydown does not always fire it */
            if (typeof editText.onChanging === "function") editText.onChanging();
        });
    }

})();
