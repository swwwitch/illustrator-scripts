#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

概要

ドキュメント内のテキストフレームから日付を検索し、選択した項目のみを置換するスクリプト。
オブジェクトが選択されている場合は、その選択範囲（グループ内を含む）のテキストフレームのみを検索対象とする。
文字の一部を選択している場合は、その親テキストフレームを検索対象に含める。

日付形式：
  検出時は、各形式に付いた曜日サフィックス（例：金曜 / 金曜日 / (金) / （金） / Fri / Friday）も対象にする。
  単独の漢字曜日（例：金）は、日付本文のサフィックスとしては検出せず、GroupItem 内の曜日のみフレームとしてのみ扱う。
  月・日は 0 を許容せず、実在しない日付（例：2026.0.8 / 2026.5.0 / 2026.2.31）は検出対象外。
  元号範囲外の日付（例：R1.1.1 ＝ 令和元年1月1日は令和開始前）も同様に対象外。
  出力時は、フォーマットドロップダウンで「元の形式を保持」または下記の形式を選択できる。
  - YYYY年M月D日 / M月D日
  - YYYY.M.D / M.D
  - YYYY/M/D / M/D
  - 令和Y年M月D日 / 平成Y年M月D日
  - RY.M.D / RY/M/D / HY.M.D / HY/M/D
  令和形式は 2019/5/1 以降、平成形式は 1989/1/8〜2019/4/30 の範囲のみ有効。
  検出時も、この範囲外の日付は対象外。

曜日表記（曜日ドロップダウン）：
  なし / 火 / 火曜 / 火曜日 / （火） / (火) / Tue / Tuesday
  ドロップダウンの表示は固定し、置換時に入力日付に応じた実際の曜日へ変換する。
  日付内の曜日サフィックス、または GroupItem 内の曜日のみフレームから初期選択を推定する。
  「元の形式を保持」でも、曜日ドロップダウンの選択を反映する。
  「なし」を選ぶと、日付内の曜日表記のみ削除する。
  連動する曜日のみフレームは、意図しない空文字化を避けるため更新しない。

数字の書式を保持：
  「元の形式を保持」を選んだ場合のみ、年・月・日それぞれの文字書式
  （フォント・サイズ・色など）をセグメント単位で維持したまま置換する。
  ドット／スラッシュ区切りの日付は、曜日サフィックス付きプレビューの安定性を優先し、
  マッチ範囲全体を1セグメントとして置換する。
  OFF、または明示フォーマット選択時は、マッチ範囲全体を先頭文字の書式に揃える。

プレビュー：
  ON で現在の入力をドキュメントに即時反映し、ダイアログを開いたまま結果を確認できる。
  入力・チェックボックス・フォーマット・曜日・数字書式保持の変更で自動更新する。
  OFF・キャンセルでは元に戻す。OK 時はプレビュー結果をいったん戻してから、本番処理として再実行する。

GroupItem 内の曜日ペアリング：
  日付フレームと、同じ GroupItem に直接含まれる「曜日のみ」のテキストフレーム
  （例：金 / 金曜 / 金曜日 / (金) / （金） / Fri / Friday）を自動で関連付ける。
  単独の漢字曜日（例：金）は、この曜日のみフレームでは有効。
  日付内に曜日がなくても、曜日のみフレームの表記スタイルを曜日ドロップダウンの初期選択に反映する。
  日付置換時には、曜日のみフレームも連動して更新する。ただし、曜日ドロップダウンが「なし」の場合は更新しない。
  同一グループ内に複数の日付フレームがある場合は、対応関係が曖昧になるため曜日フレームをペアリングしない。

入力補助：
  年・月・日の入力欄では、↑↓キーで±1、Shift＋↑↓で±10増減する。
  0 以下にはならない。

検索結果はアートボードごとにグループ化して表示する。

注意点：
  - GroupItem 内の曜日ペアリングは「直接の子」のみが対象。ネストした GroupItem 内の
    曜日フレームは検出されない。必要なら、日付フレームと同じ階層に置くか、
    グループ構造を平坦化する。
  - プレビューの巻き戻しは app.undo() を文字操作回数ぶん呼ぶ実装。
    Illustrator の undo グループ化の挙動に依存するため、
    環境や操作状況によっては undo 回数がずれる可能性がある。
    プレビューを多用したあとに状態が不安定になった場合は、
    キャンセルで戻すか、ダイアログを開き直す。
  - 「令和Y.M.D」「平成Y.M.D」（ドット区切りの和暦）は検出対象外。
    ドット／スラッシュ区切りの略記は RY.M.D / RY/M/D / HY.M.D / HY/M/D を使用する。

*/

(function () {

    // =========================================
    // バージョン
    // =========================================

    var SCRIPT_VERSION = "v1.0.2";

    /* エラー件数メッセージ */
    function formatErrorMessage(count) {
        return count + "件はロック等の理由で処理できませんでした。";
    }

    // =========================================
    // 前提チェック
    // =========================================

    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;

    // =========================================
    // 共通ヘルパー
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10];

    /* パネルの共通設定 */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            panel.spacing = spacing;
        }
    }

    /* テキストフレームが属するアートボードを判定 */
    function getArtboardIndex(item) {
        var bounds = item.geometricBounds;
        var centerX = (bounds[0] + bounds[2]) / 2;
        var centerY = (bounds[1] + bounds[3]) / 2;

        for (var artboardIdx = 0; artboardIdx < doc.artboards.length; artboardIdx++) {
            var rect = doc.artboards[artboardIdx].artboardRect;
            var minX = Math.min(rect[0], rect[2]);
            var maxX = Math.max(rect[0], rect[2]);
            var minY = Math.min(rect[1], rect[3]);
            var maxY = Math.max(rect[1], rect[3]);
            if (centerX >= minX && centerX <= maxX && centerY >= minY && centerY <= maxY) {
                return artboardIdx;
            }
        }
        return -1;
    }

    /* ↑↓キーで値を増減（Shiftで±10） */
    function changeValueByArrowKey(editText, onValueChanged) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
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

    /* 一致した文字列の曜日サフィックスの形式を判定 */
    function detectWeekdaySuffixStyle(matchText) {
        var parsed = detectFormatAndParse(matchText);
        if (!parsed || !parsed.parts || !parsed.parts.suffix) return null;
        var suffix = parsed.parts.suffix;
        if (/^[日月火水木金土]曜日$/.test(suffix)) return 'long';           /* 例：金曜日 */
        if (/^[日月火水木金土]曜$/.test(suffix)) return 'medium';           /* 例：金曜 */
        if (/^\([日月火水木金土]\)$/.test(suffix)) return 'half-paren';   /* 例：(金) */
        if (/^（[日月火水木金土]）$/.test(suffix)) return 'full-paren';    /* 例：（金） */
        if (/^(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat)day$/.test(suffix)) return 'en-full';
        if (/^(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat)$/.test(suffix)) return 'en-short';
        return null;
    }

    /* 令和元年 = 2019 年（令和Y = 西暦 Y + 2018）／平成元年 = 1989 年（平成Y = 西暦 Y + 1988） */
    var REIWA_BASE_YEAR = 2018;
    var HEISEI_BASE_YEAR = 1988;

    /* 元号→西暦／西暦→元号 */
    function reiwaToGregorian(reiwaYear) { return reiwaYear + REIWA_BASE_YEAR; }
    function gregorianToReiwa(gregorianYear) { return gregorianYear - REIWA_BASE_YEAR; }
    function heiseiToGregorian(heiseiYear) { return heiseiYear + HEISEI_BASE_YEAR; }
    function gregorianToHeisei(gregorianYear) { return gregorianYear - HEISEI_BASE_YEAR; }

    /* 元号年の文字列化。漢字フォーマット（reiwa-jp / heisei-jp）では元年 1 を「元」と表記する */
    function formatEraYearText(eraYear, useGanText) {
        if (useGanText && eraYear === 1) return "元";
        return String(eraYear);
    }

    /* 元の数字テキストが 2 桁ゼロ詰め（"05" 等）なら、新しい数字も同じ桁数に揃える。
       新しい文字列が数字以外（例：「元」）の場合は揃えない */
    function applyZeroPaddingFrom(oldText, newValue) {
        var newStr = String(newValue);
        if (oldText.length === 2 && oldText.charAt(0) === '0' && newStr.length === 1 && /^[0-9]$/.test(newStr)) {
            return "0" + newStr;
        }
        return newStr;
    }

    /* 各元号の有効範囲（getValidationErrorMessage で使用） */
    var REIWA_START_DATE = new Date(2019, 4, 1);                 /* 2019/5/1 〜 */
    var HEISEI_START_DATE = new Date(1989, 0, 8);                /* 1989/1/8 〜 */
    var HEISEI_END_DATE_EXCLUSIVE = new Date(2019, 4, 1);        /* 〜 2019/4/30 */

    /* 実在する日付かを判定 / Check whether the date exists */
    function isRealDate(year, month, day) {
        var checkDate = new Date(year, month - 1, day);
        return checkDate.getFullYear() === year &&
            checkDate.getMonth() === month - 1 &&
            checkDate.getDate() === day;
    }

    /* 元号形式の日付が元号の有効範囲内かを判定 / Check whether an era date is within its valid range */
    function isEraDateValid(parsed) {
        if (!parsed) return false;
        var checkDate = new Date(parsed.year, parsed.month - 1, parsed.day);
        if (parsed.era === 'reiwa') return checkDate >= REIWA_START_DATE;
        if (parsed.era === 'heisei') return checkDate >= HEISEI_START_DATE && checkDate < HEISEI_END_DATE_EXCLUSIVE;
        return true;
    }

    /*
       マッチ文字列を解析し、形式種別と各構成要素を返す。
       戻り値の主要フィールド：
         format: 'jp' | 'dot' | 'slash' | 'reiwa-jp' | 'r-dot' | 'r-slash' | 'heisei-jp' | 'h-dot' | 'h-slash'
         era:    'reiwa' | 'heisei' | null
         year, month, day: 西暦の数値（令和・平成形式の場合は変換後）
         parts:  元テキストの分解（prefix, year, sep1, month, sep2, day, sep3, suffix）
    */
    function detectFormatAndParse(matchText) {
        var m;

        /* 令和Y年M月D日（曜日サフィックスは末尾に残す） */
        m = String(matchText).match(/^(令和)([0-9]{1,2})(年)(0?[1-9]|1[0-2])(月)(0?[1-9]|[12][0-9]|3[01])(日)(.*)$/);
        if (m) {
            return {
                format: 'reiwa-jp', era: 'reiwa',
                year: reiwaToGregorian(parseInt(m[2], 10)),
                month: parseInt(m[4], 10),
                day: parseInt(m[6], 10),
                parts: { prefix: m[1], year: m[2], sep1: m[3], month: m[4], sep2: m[5], day: m[6], sep3: m[7], suffix: m[8] || "" }
            };
        }

        /* 平成Y年M月D日（曜日サフィックスは末尾に残す） */
        m = String(matchText).match(/^(平成)([0-9]{1,2})(年)(0?[1-9]|1[0-2])(月)(0?[1-9]|[12][0-9]|3[01])(日)(.*)$/);
        if (m) {
            return {
                format: 'heisei-jp', era: 'heisei',
                year: heiseiToGregorian(parseInt(m[2], 10)),
                month: parseInt(m[4], 10),
                day: parseInt(m[6], 10),
                parts: { prefix: m[1], year: m[2], sep1: m[3], month: m[4], sep2: m[5], day: m[6], sep3: m[7], suffix: m[8] || "" }
            };
        }

        /* RY.M.D（曜日サフィックスは末尾に残す） */
        m = String(matchText).match(/^(R)([0-9]{1,2})\.(0?[1-9]|1[0-2])\.(0?[1-9]|[12][0-9]|3[01])(.*)$/);
        if (m) {
            return {
                format: 'r-dot', era: 'reiwa',
                year: reiwaToGregorian(parseInt(m[2], 10)),
                month: parseInt(m[3], 10),
                day: parseInt(m[4], 10),
                parts: { prefix: m[1], year: m[2], sep1: ".", month: m[3], sep2: ".", day: m[4], sep3: "", suffix: m[5] || "" }
            };
        }

        /* RY/M/D（曜日サフィックスは末尾に残す） */
        m = String(matchText).match(/^(R)([0-9]{1,2})\/(0?[1-9]|1[0-2])\/(0?[1-9]|[12][0-9]|3[01])(.*)$/);
        if (m) {
            return {
                format: 'r-slash', era: 'reiwa',
                year: reiwaToGregorian(parseInt(m[2], 10)),
                month: parseInt(m[3], 10),
                day: parseInt(m[4], 10),
                parts: { prefix: m[1], year: m[2], sep1: "/", month: m[3], sep2: "/", day: m[4], sep3: "", suffix: m[5] || "" }
            };
        }

        /* HY.M.D（曜日サフィックスは末尾に残す） */
        m = String(matchText).match(/^(H)([0-9]{1,2})\.(0?[1-9]|1[0-2])\.(0?[1-9]|[12][0-9]|3[01])(.*)$/);
        if (m) {
            return {
                format: 'h-dot', era: 'heisei',
                year: heiseiToGregorian(parseInt(m[2], 10)),
                month: parseInt(m[3], 10),
                day: parseInt(m[4], 10),
                parts: { prefix: m[1], year: m[2], sep1: ".", month: m[3], sep2: ".", day: m[4], sep3: "", suffix: m[5] || "" }
            };
        }

        /* HY/M/D（曜日サフィックスは末尾に残す） */
        m = String(matchText).match(/^(H)([0-9]{1,2})\/(0?[1-9]|1[0-2])\/(0?[1-9]|[12][0-9]|3[01])(.*)$/);
        if (m) {
            return {
                format: 'h-slash', era: 'heisei',
                year: heiseiToGregorian(parseInt(m[2], 10)),
                month: parseInt(m[3], 10),
                day: parseInt(m[4], 10),
                parts: { prefix: m[1], year: m[2], sep1: "/", month: m[3], sep2: "/", day: m[4], sep3: "", suffix: m[5] || "" }
            };
        }

        /* YYYY年M月D日（曜日サフィックスは末尾に残す） */
        m = String(matchText).match(/^([0-9]{4})(年)(0?[1-9]|1[0-2])(月)(0?[1-9]|[12][0-9]|3[01])(日)(.*)$/);
        if (m) {
            return {
                format: 'jp', era: null,
                year: parseInt(m[1], 10),
                month: parseInt(m[3], 10),
                day: parseInt(m[5], 10),
                parts: { prefix: "", year: m[1], sep1: m[2], month: m[3], sep2: m[4], day: m[5], sep3: m[6], suffix: m[7] || "" }
            };
        }

        /* YYYY.M.D（曜日サフィックスは末尾に残す） */
        m = String(matchText).match(/^([0-9]{4})\.(0?[1-9]|1[0-2])\.(0?[1-9]|[12][0-9]|3[01])(.*)$/);
        if (m) {
            return {
                format: 'dot', era: null,
                year: parseInt(m[1], 10),
                month: parseInt(m[2], 10),
                day: parseInt(m[3], 10),
                parts: { prefix: "", year: m[1], sep1: ".", month: m[2], sep2: ".", day: m[3], sep3: "", suffix: m[4] || "" }
            };
        }

        /* YYYY/M/D（曜日サフィックスは末尾に残す） */
        m = String(matchText).match(/^([0-9]{4})\/(0?[1-9]|1[0-2])\/(0?[1-9]|[12][0-9]|3[01])(.*)$/);
        if (m) {
            return {
                format: 'slash', era: null,
                year: parseInt(m[1], 10),
                month: parseInt(m[2], 10),
                day: parseInt(m[3], 10),
                parts: { prefix: "", year: m[1], sep1: "/", month: m[2], sep2: "/", day: m[3], sep3: "", suffix: m[4] || "" }
            };
        }

        return null;
    }


    /* テキストフレーム内の一致範囲を書式を保持したまま置換。実行した文字操作数を返す（プレビュー巻き戻し用） */
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

    /* 2 つの日付の差を日数で返す */
    function getDaysDifference(fromDate, toDate) {
        var msPerDay = 1000 * 60 * 60 * 24;
        return Math.round((toDate.getTime() - fromDate.getTime()) / msPerDay);
    }

    /* 日数差ラベル（符号付き、例：「+15日」） */
    function formatDaysDifference(days) {
        var sign = days > 0 ? "+" : "";
        return sign + days + "日";
    }

    /* 曜日表示（例：金曜日） */
    function getWeekdayLabel(year, month, day) {
        if (isNaN(year) || isNaN(month) || isNaN(day)) return "";
        var checkDate = new Date(year, month - 1, day);
        if (isNaN(checkDate.getTime())) return "";
        var jaDays = ["日", "月", "火", "水", "木", "金", "土"];
        return jaDays[checkDate.getDay()] + "曜日";
    }

    /* 西暦+月日から和暦表記を返す（例：「令和8年」「平成元年」）。明治より前は空文字 */
    function formatEraLabel(year, month, day) {
        if (isNaN(year) || isNaN(month) || isNaN(day)) return "";
        var checkDate = new Date(year, month - 1, day);
        if (isNaN(checkDate.getTime())) return "";

        var era, eraYear;
        if (checkDate >= new Date(2019, 4, 1)) {       /* 令和：2019/5/1〜 */
            era = "令和"; eraYear = year - 2018;
        } else if (checkDate >= new Date(1989, 0, 8)) { /* 平成：1989/1/8〜2019/4/30 */
            era = "平成"; eraYear = year - 1988;
        } else if (checkDate >= new Date(1926, 11, 25)) { /* 昭和：1926/12/25〜1989/1/7 */
            era = "昭和"; eraYear = year - 1925;
        } else if (checkDate >= new Date(1912, 6, 30)) { /* 大正：1912/7/30〜1926/12/24 */
            era = "大正"; eraYear = year - 1911;
        } else if (checkDate >= new Date(1868, 8, 8)) {  /* 明治：1868/9/8〜1912/7/29 */
            era = "明治"; eraYear = year - 1867;
        } else {
            return "";
        }

        var yearText = (eraYear === 1) ? "元" : String(eraYear);
        return era + yearText + "年";
    }

    // =========================================
    // 日付検索
    // =========================================

    /*
       対象例：
         2026年5月8日 / 2026年5月8日金曜 / 2026年5月8日金曜日 / 2026年5月8日(金) / 2026年5月8日（金）
         2026.5.8 / 2026.5.8（金） / 2026/5/8 / 2026/5/8（金）
         令和8年5月8日 / 令和8年5月8日（金） / R8.5.8 / R8.5.8（金） / R8/5/8 / R8/5/8（金）
         英語曜日サフィックス（Fri / Friday など）も検出する。
       元号系を先に列挙して優先マッチさせる。
    */

    var monthPattern = "(?:0?[1-9]|1[0-2])";
    var dayPattern = "(?:0?[1-9]|[12][0-9]|3[01])";

    var weekdaySuffixPattern = "(?:[日月火水木金土]曜日|[日月火水木金土]曜|\\([日月火水木金土]\\)|（[日月火水木金土]）|Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sun|Mon|Tue|Wed|Thu|Fri|Sat)?";
    var dateRegex = new RegExp(
        "令和[0-9]{1,2}年" + monthPattern + "月" + dayPattern + "日" + weekdaySuffixPattern +
        "|平成[0-9]{1,2}年" + monthPattern + "月" + dayPattern + "日" + weekdaySuffixPattern +
        "|R[0-9]{1,2}\\." + monthPattern + "\\." + dayPattern + weekdaySuffixPattern +
        "|R[0-9]{1,2}/" + monthPattern + "/" + dayPattern + weekdaySuffixPattern +
        "|H[0-9]{1,2}\\." + monthPattern + "\\." + dayPattern + weekdaySuffixPattern +
        "|H[0-9]{1,2}/" + monthPattern + "/" + dayPattern + weekdaySuffixPattern +
        "|[0-9]{4}年" + monthPattern + "月" + dayPattern + "日" + weekdaySuffixPattern +
        "|[0-9]{4}\\." + monthPattern + "\\." + dayPattern + weekdaySuffixPattern +
        "|[0-9]{4}/" + monthPattern + "/" + dayPattern + weekdaySuffixPattern,
        "g"
    );

    /* テキストフレームを重複なく追加 / Add a text frame without duplicates */
    function addTextFrameOnce(textFrame, accumulator) {
        if (!textFrame || textFrame.typename !== "TextFrame") return;
        for (var frameIndex = 0; frameIndex < accumulator.length; frameIndex++) {
            if (accumulator[frameIndex] === textFrame) return;
        }
        accumulator.push(textFrame);
    }

    /* 部分テキスト選択から親テキストフレームを取得 / Get parent text frame from partial text selection */
    function getParentTextFrameFromTextSelection(item) {
        var currentItem = item;
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

    /* 選択オブジェクト（およびその子）からテキストフレームを再帰的に収集 */
    function collectTextFramesFromItems(items, accumulator) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) continue;
            if (item.typename === "TextFrame") {
                addTextFrameOnce(item, accumulator);
            } else if (item.typename === "GroupItem") {
                collectTextFramesFromItems(item.pageItems, accumulator);
            } else if (item.typename === "TextRange" || item.typename === "InsertionPoint" || item.typename === "Character") {
                addTextFrameOnce(getParentTextFrameFromTextSelection(item), accumulator);
            }
        }
    }

    var foundMatches = [];
    var textFrames = [];
    var selection = doc.selection;

    if (selection && selection.length > 0) {
        /* 選択あり：選択範囲内のテキストフレームのみ */
        collectTextFramesFromItems(selection, textFrames);
    } else {
        /* 選択なし：ドキュメント全体 */
        for (var docFrameIdx = 0; docFrameIdx < doc.textFrames.length; docFrameIdx++) {
            textFrames.push(doc.textFrames[docFrameIdx]);
        }
    }

    for (var frameIdx = 0; frameIdx < textFrames.length; frameIdx++) {
        var frame = textFrames[frameIdx];
        var frameContent = "";

        try {
            frameContent = frame.contents;
        } catch (eRead) {
            continue;
        }

        var artboardIdx = -1;
        try {
            artboardIdx = getArtboardIndex(frame);
        } catch (eArtboard) {
            artboardIdx = -1;
        }

        var match;
        dateRegex.lastIndex = 0;

        while ((match = dateRegex.exec(frameContent)) !== null) {
            /* 直前が数字、かつマッチが数字始まり（YYYY 系）の場合は、5 桁以上の連続数字を切り出した
               誤検出とみなしてスキップ。R/H/令和/平成 始まりはこの判定の対象外 */
            if (match.index > 0 &&
                /[0-9]/.test(frameContent.charAt(match.index - 1)) &&
                /^[0-9]/.test(match[0])) {
                continue;
            }
            var parsed = detectFormatAndParse(match[0]);
            if (!parsed) continue;
            if (!isRealDate(parsed.year, parsed.month, parsed.day)) continue;
            if (!isEraDateValid(parsed)) continue;
            foundMatches.push({
                frameIndex: frameIdx,
                text: match[0],
                matchIndex: match.index,
                artboardIndex: artboardIdx,
                parsed: parsed,
                weekdaySuffixStyle: detectWeekdaySuffixStyle(match[0]),
                weekdayPairs: null,
                checkbox: null
            });
        }
    }

    // =========================================
    // 同一グループ内の曜日フレームをペアリング
    // =========================================

    /* フレーム全体が曜日のみかを判定し、スタイル名を返す（WEEKDAY_VALUES と同じ命名）。単独漢字曜日はここでのみ許可 */
    function detectWeekdayOnlyFrame(content) {
        var trimmed = String(content).replace(/^\s+|\s+$/g, "");
        if (/^[日月火水木金土]曜日$/.test(trimmed)) return 'long';
        if (/^[日月火水木金土]曜$/.test(trimmed)) return 'medium';
        if (/^\([日月火水木金土]\)$/.test(trimmed)) return 'half-paren';
        if (/^（[日月火水木金土]）$/.test(trimmed)) return 'full-paren';
        if (/^[日月火水木金土]$/.test(trimmed)) return 'kanji';
        if (/^(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat)day$/.test(trimmed)) return 'en-full';
        if (/^(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat)$/.test(trimmed)) return 'en-short';
        return null;
    }

    /*
       日付マッチごとに、その TextFrame の直接の親 GroupItem を見て、
       同じ親グループ内の「曜日のみ」TextFrame を曜日ペアとして関連付ける。
       同一グループに複数の日付マッチがある場合は対応が曖昧なのでペアリングしない。
    */
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

    // =========================================
    // 日付の初期値
    // =========================================

    var today = new Date();
    var defaultYear = today.getFullYear();
    var defaultMonth = today.getMonth() + 1;
    var defaultDay = today.getDate();

    // =========================================
    // ダイアログ構築
    // =========================================

    var dialog = new Window("dialog", "日付を検索・置換 " + SCRIPT_VERSION);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";

    var allCheckboxes = [];

    /* チェックボックスに Option＋クリックで全切替を割り当てる。クリック後はプレビューも更新 */
    function bindOptionClickToggleAll(checkboxControl) {
        checkboxControl.addEventListener("click", function (event) {
            if (event.altKey) {
                var newValue = checkboxControl.value;
                for (var idx = 0; idx < allCheckboxes.length; idx++) {
                    allCheckboxes[idx].value = newValue;
                }
            }
            refreshPreview();
        });
    }

    if (foundMatches.length === 0) {
        dialog.add("statictext", undefined, "日付は見つかりませんでした。");
    } else {
        /* アートボードごとに結果をグループ化 */
        var matchesByArtboard = {};
        var orderedArtboardKeys = [];
        for (var matchIdx = 0; matchIdx < foundMatches.length; matchIdx++) {
            var artboardKey = String(foundMatches[matchIdx].artboardIndex);
            if (!matchesByArtboard[artboardKey]) {
                matchesByArtboard[artboardKey] = [];
                orderedArtboardKeys.push(foundMatches[matchIdx].artboardIndex);
            }
            matchesByArtboard[artboardKey].push(foundMatches[matchIdx]);
        }
        /* -1（アートボード外）は末尾に */
        orderedArtboardKeys.sort(function (a, b) {
            if (a === -1 && b !== -1) return 1;
            if (b === -1 && a !== -1) return -1;
            return a - b;
        });

        var foundDatesPanel = dialog.add("panel", undefined, "見つかった日付（" + foundMatches.length + "）");
        setupPanel(foundDatesPanel, 6);
        foundDatesPanel.minimumSize.width = 240;

        for (var keyIdx = 0; keyIdx < orderedArtboardKeys.length; keyIdx++) {
            var currentArtboardIdx = orderedArtboardKeys[keyIdx];
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
            var artboardHeader = artboardHeaderGroup.add("statictext", undefined, artboardHeaderText);

            var artboardItems = matchesByArtboard[String(currentArtboardIdx)];
            for (var itemIdx = 0; itemIdx < artboardItems.length; itemIdx++) {
                /* 各日付行は左マージン 20px でインデント */
                var checkboxRow = foundDatesPanel.add("group");
                checkboxRow.orientation = "row";
                checkboxRow.alignment = "left";
                checkboxRow.margins = [20, 0, 0, 0];
                var labelText = artboardItems[itemIdx].text;
                var pairCount = (artboardItems[itemIdx].weekdayPairs ? artboardItems[itemIdx].weekdayPairs.length : 0);
                if (pairCount > 0) {
                    labelText += "  ＋曜日連動";
                }
                var checkbox = checkboxRow.add("checkbox", undefined, labelText);
                checkbox.value = true;
                /* チェックボックスのラベル切れ対策 */
                checkbox.preferredSize.width = 220;
                checkbox.helpTip = (pairCount > 0)
                    ? ("Option＋クリックで全項目を一括切替\n同一グループ内の曜日フレーム " + pairCount + " 件も連動して更新されます\n曜日が「なし」の場合は連動フレームを更新しません")
                    : "Option＋クリックで全項目を一括切替";
                artboardItems[itemIdx].checkbox = checkbox;
                allCheckboxes.push(checkbox);
                bindOptionClickToggleAll(checkbox);
            }
        }
    }

    /* 置換後の日付入力パネル：［2026］年［5］月［8］日 形式 */
    var replacementPanel = dialog.add("panel", undefined, "置換後の日付");
    setupPanel(replacementPanel, 6);

    var dateInputGroup = replacementPanel.add("group");
    dateInputGroup.orientation = "row";

    var ARROW_KEY_HELP = "↑↓キーで増減（Shift：±10）";

    var yearInput = dateInputGroup.add("edittext", undefined, String(defaultYear));
    yearInput.characters = 5;
    yearInput.helpTip = ARROW_KEY_HELP;
    dateInputGroup.add("statictext", undefined, "年");

    var monthInput = dateInputGroup.add("edittext", undefined, String(defaultMonth));
    monthInput.characters = 3;
    monthInput.helpTip = ARROW_KEY_HELP;
    dateInputGroup.add("statictext", undefined, "月");

    var dayInput = dateInputGroup.add("edittext", undefined, String(defaultDay));
    dayInput.characters = 3;
    dayInput.helpTip = ARROW_KEY_HELP;
    dateInputGroup.add("statictext", undefined, "日");

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

    var formatRow = replacementPanel.add("group");
    formatRow.orientation = "row";
    formatRow.add("statictext", undefined, "フォーマット：");
    var formatDropdown = formatRow.add("dropdownlist", undefined, FORMAT_LABELS);
    formatDropdown.selection = 0;
    formatDropdown.helpTip = "「元の形式を保持」：各マッチの区切り文字・元号を維持\n曜日表記は右の曜日ドロップダウンの選択を反映\nそれ以外：すべてのマッチを選択した形式に統一";

    var weekdayDropdown = formatRow.add("dropdownlist", undefined, WEEKDAY_LABELS);
    weekdayDropdown.helpTip = "出力に付与する曜日表記。「元の形式を保持」選択時もこの選択を反映し、「なし」で日付内の曜日表記を削除。連動曜日フレームは更新しません";

    /* 最初に見つかった置換対象を基準に、曜日サフィックスまたは曜日のみフレームの形式を初期選択にする。
       日付サフィックスは単独漢字曜日を許可しないため 'kanji' は返らない。曜日のみフレームでのみ拾う */
    function getSuffixStyleAsWeekdayChoice(matchInfo) {
        if (!matchInfo) return null;
        var suffixStyle = matchInfo.weekdaySuffixStyle;
        if (suffixStyle === 'medium') return 'medium';
        if (suffixStyle === 'long') return 'long';
        if (suffixStyle === 'full-paren') return 'full-paren';
        if (suffixStyle === 'half-paren') return 'half-paren';
        if (suffixStyle === 'en-short') return 'en-short';
        if (suffixStyle === 'en-full') return 'en-full';
        return null;
    }

    function getPairStyleAsWeekdayChoice(matchInfo) {
        if (!matchInfo || !matchInfo.weekdayPairs || matchInfo.weekdayPairs.length === 0) return null;

        for (var pairIndex = 0; pairIndex < matchInfo.weekdayPairs.length; pairIndex++) {
            var pairStyle = matchInfo.weekdayPairs[pairIndex].style;
            if (pairStyle === 'kanji') return 'kanji';
            if (pairStyle === 'medium') return 'medium';
            if (pairStyle === 'long') return 'long';
            if (pairStyle === 'full-paren') return 'full-paren';
            if (pairStyle === 'half-paren') return 'half-paren';
            if (pairStyle === 'en-short') return 'en-short';
            if (pairStyle === 'en-full') return 'en-full';
        }

        return null;
    }

    function detectInitialWeekdayChoice() {
        var choice;

        for (var matchIndex = 0; matchIndex < foundMatches.length; matchIndex++) {
            choice = getSuffixStyleAsWeekdayChoice(foundMatches[matchIndex]);
            if (choice) return choice;

            choice = getPairStyleAsWeekdayChoice(foundMatches[matchIndex]);
            if (choice) return choice;
        }

        return 'none';
    }

    var initialWeekdayChoice = detectInitialWeekdayChoice();
    var initialWeekdayIndex = 0;
    for (var wi = 0; wi < WEEKDAY_VALUES.length; wi++) {
        if (WEEKDAY_VALUES[wi] === initialWeekdayChoice) { initialWeekdayIndex = wi; break; }
    }
    weekdayDropdown.selection = initialWeekdayIndex;
    var currentWeekdayChoice = WEEKDAY_VALUES[initialWeekdayIndex];

    /* 曜日ドロップダウンの現在値を内部値として保持する */
    function updateCurrentWeekdayChoice() {
        try {
            if (weekdayDropdown.selection && typeof weekdayDropdown.selection.index === 'number') {
                currentWeekdayChoice = WEEKDAY_VALUES[weekdayDropdown.selection.index];
            }
        } catch (eWeekdayChoice) { }
    }


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
    var checkPanel = dialog.add("panel", undefined, "置換内容の確認");
    setupPanel(checkPanel, 6);

    /* ラベル幅を揃えて値の左端を縦に整列 */
    var CHECK_LABEL_WIDTH = 130;

    function addCheckRow(labelText, valueWidth) {
        var row = checkPanel.add("group");
        row.orientation = "row";
        var head = row.add("statictext", undefined, labelText);
        head.preferredSize.width = CHECK_LABEL_WIDTH;
        var value = row.add("statictext", undefined, "");
        value.preferredSize.width = valueWidth;
        return value;
    }

    var eraLabel = addCheckRow("和暦：", 120);
    var weekdayLabel = addCheckRow("置換後の曜日：", 80);
    var daysDiffLabel = addCheckRow("最初の日付との差：", 120);

    /* 比較用の基準日（最初に見つかった日付。形式に関わらず西暦で扱う） */
    var referenceDate = null;
    if (foundMatches.length > 0 && foundMatches[0].parsed) {
        var refParsed = foundMatches[0].parsed;
        referenceDate = new Date(refParsed.year, refParsed.month - 1, refParsed.day);
    }

    /* 入力値からチェック用の表示を更新。プレビュー ON 時は再適用も走らせる */
    function refreshCheckLabels() {
        var y = parseInt(yearInput.text, 10);
        var m = parseInt(monthInput.text, 10);
        var d = parseInt(dayInput.text, 10);

        eraLabel.text = formatEraLabel(y, m, d);
        weekdayLabel.text = getWeekdayLabel(y, m, d);

        if (referenceDate && !isNaN(y) && !isNaN(m) && !isNaN(d)) {
            var newDate = new Date(y, m - 1, d);
            if (!isNaN(newDate.getTime())) {
                daysDiffLabel.text = formatDaysDifference(getDaysDifference(referenceDate, newDate));
            } else {
                daysDiffLabel.text = "";
            }
        } else {
            daysDiffLabel.text = "";
        }

        refreshPreview();
    }

    yearInput.onChange = refreshCheckLabels;
    monthInput.onChange = refreshCheckLabels;
    dayInput.onChange = refreshCheckLabels;

    changeValueByArrowKey(yearInput, refreshCheckLabels);
    changeValueByArrowKey(monthInput, refreshCheckLabels);
    changeValueByArrowKey(dayInput, refreshCheckLabels);

    refreshCheckLabels();

    // =========================================
    // プレビュー / 実行
    // =========================================

    var isPreviewApplied = false;
    var previewOpCount = 0;

    /* 入力値の検証。エラーがあればメッセージ文字列、なければ null を返す */
    function getValidationErrorMessage() {
        var y = parseInt(yearInput.text, 10);
        var m = parseInt(monthInput.text, 10);
        var d = parseInt(dayInput.text, 10);
        if (isNaN(y) || isNaN(m) || isNaN(d)) return "年・月・日は半角数字で入力してください。";
        if (y < 1) return "年は1以上で入力してください。";
        if (m < 1 || m > 12) return "月は1〜12で入力してください。";
        if (d < 1 || d > 31) return "日は1〜31で入力してください。";
        var checkDate = new Date(y, m - 1, d);
        if (
            checkDate.getFullYear() !== y ||
            checkDate.getMonth() !== m - 1 ||
            checkDate.getDate() !== d
        ) return "存在しない日付です。";

        /* 元号フォーマット選択時は、各元号の有効範囲に収まることを要求 */
        var formatChoice = getDropdownValue(formatDropdown, FORMAT_VALUES, 'preserve');
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
                var fm = foundMatches[fmIdx];
                if (fm.checkbox && !fm.checkbox.value) continue;
                if (!fm.parsed) continue;
                if (fm.parsed.era === 'reiwa') hasReiwaMatch = true;
                if (fm.parsed.era === 'heisei') hasHeiseiMatch = true;
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

    function isInputValid() {
        return getValidationErrorMessage() === null;
    }

    function buildExplicitWeekday(weekdayChoice, year, month, day) {
        if (weekdayChoice === 'none') return "";
        var checkDate = new Date(year, month - 1, day);
        if (isNaN(checkDate.getTime())) return "";
        var idx = checkDate.getDay();
        switch (weekdayChoice) {
            case 'kanji': return WEEKDAY_KANJI[idx];
            case 'medium': return WEEKDAY_KANJI[idx] + "曜";
            case 'long': return WEEKDAY_KANJI[idx] + "曜日";
            case 'full-paren': return "（" + WEEKDAY_KANJI[idx] + "）";
            case 'half-paren': return "(" + WEEKDAY_KANJI[idx] + ")";
            case 'en-short': return WEEKDAY_EN_SHORT[idx];
            case 'en-full': return WEEKDAY_EN_FULL[idx];
        }
        return "";
    }

    /* dropdownlist の選択を安全に取得（refresh 直後に selection が外れているケースに備える） */
    function getDropdownValue(dropdown, valuesArray, fallback) {
        try {
            if (dropdown.selection !== null && dropdown.selection !== undefined) {
                var idx = dropdown.selection.index;
                if (typeof idx === 'number' && idx >= 0 && idx < valuesArray.length) {
                    return valuesArray[idx];
                }
            }
        } catch (eDD) { }
        return fallback;
    }

    /* チェックされたマッチに対して置換を実行。文字操作数・エラー件数・対象件数を返す */
    function performReplacement() {
        var newYear = parseInt(yearInput.text, 10);
        var newMonth = parseInt(monthInput.text, 10);
        var newDay = parseInt(dayInput.text, 10);
        var formatChoice = getDropdownValue(formatDropdown, FORMAT_VALUES, 'preserve');
        var weekdayChoice = currentWeekdayChoice || 'none';
        var preserveNumberFormat = preserveNumberFormatCheckbox.value;
        var explicitWeekdayText = buildExplicitWeekday(weekdayChoice, newYear, newMonth, newDay);

        function buildPreservedFormatText(matchInfo) {
            var parsed = matchInfo.parsed;
            if (!parsed) return null;
            var oldParts = parsed.parts;
            var newYearText;
            if (parsed.era === 'reiwa') newYearText = formatEraYearText(gregorianToReiwa(newYear), parsed.format === 'reiwa-jp');
            else if (parsed.era === 'heisei') newYearText = formatEraYearText(gregorianToHeisei(newYear), parsed.format === 'heisei-jp');
            else newYearText = String(newYear);
            newYearText = applyZeroPaddingFrom(oldParts.year, newYearText);
            var text =
                oldParts.prefix +
                newYearText +
                oldParts.sep1 +
                applyZeroPaddingFrom(oldParts.month, newMonth) +
                oldParts.sep2 +
                applyZeroPaddingFrom(oldParts.day, newDay) +
                oldParts.sep3;
            /* 曜日ドロップダウンの選択を常に反映（「なし」なら元の suffix を削除） */
            text += explicitWeekdayText;
            return text;
        }

        function buildExplicitFormatText(format) {
            var reiwaYear = gregorianToReiwa(newYear);
            var heiseiYear = gregorianToHeisei(newYear);
            var reiwaJpYearText = formatEraYearText(reiwaYear, true);
            var heiseiJpYearText = formatEraYearText(heiseiYear, true);
            switch (format) {
                case 'jp': return newYear + "年" + newMonth + "月" + newDay + "日" + explicitWeekdayText;
                case 'jp-md': return newMonth + "月" + newDay + "日" + explicitWeekdayText;
                case 'dot': return newYear + "." + newMonth + "." + newDay + explicitWeekdayText;
                case 'dot-md': return newMonth + "." + newDay + explicitWeekdayText;
                case 'slash': return newYear + "/" + newMonth + "/" + newDay + explicitWeekdayText;
                case 'slash-md': return newMonth + "/" + newDay + explicitWeekdayText;
                case 'reiwa-jp': return "令和" + reiwaJpYearText + "年" + newMonth + "月" + newDay + "日" + explicitWeekdayText;
                case 'r-dot': return "R" + reiwaYear + "." + newMonth + "." + newDay + explicitWeekdayText;
                case 'r-slash': return "R" + reiwaYear + "/" + newMonth + "/" + newDay + explicitWeekdayText;
                case 'heisei-jp': return "平成" + heiseiJpYearText + "年" + newMonth + "月" + newDay + "日" + explicitWeekdayText;
                case 'h-dot': return "H" + heiseiYear + "." + newMonth + "." + newDay + explicitWeekdayText;
                case 'h-slash': return "H" + heiseiYear + "/" + newMonth + "/" + newDay + explicitWeekdayText;
            }
            return null;
        }

        function buildReplacementText(matchInfo) {
            if (formatChoice === 'preserve') return buildPreservedFormatText(matchInfo);
            return buildExplicitFormatText(formatChoice);
        }

        function buildReplacementSegments(matchInfo) {
            var parsed = matchInfo.parsed;
            if (!parsed) return null;
            var oldParts = parsed.parts;
            var segments = [];
            var pos = 0;
            function pushSegment(oldText, newText) {
                if (oldText.length === 0) return;
                segments.push({ offset: pos, oldLen: oldText.length, newText: newText });
                pos += oldText.length;
            }
            var newYearText;
            if (parsed.era === 'reiwa') newYearText = formatEraYearText(gregorianToReiwa(newYear), parsed.format === 'reiwa-jp');
            else if (parsed.era === 'heisei') newYearText = formatEraYearText(gregorianToHeisei(newYear), parsed.format === 'heisei-jp');
            else newYearText = String(newYear);
            newYearText = applyZeroPaddingFrom(oldParts.year, newYearText);
            var hasJpSuffixSlot = (parsed.format === 'jp' || parsed.format === 'reiwa-jp' || parsed.format === 'heisei-jp');

            pushSegment(oldParts.prefix, oldParts.prefix);
            pushSegment(oldParts.year, newYearText);
            pushSegment(oldParts.sep1, oldParts.sep1);
            pushSegment(oldParts.month, applyZeroPaddingFrom(oldParts.month, newMonth));
            pushSegment(oldParts.sep2, oldParts.sep2);

            if (hasJpSuffixSlot) {
                pushSegment(oldParts.day, applyZeroPaddingFrom(oldParts.day, newDay));
                if (oldParts.suffix.length > 0) {
                    /* 元 suffix を曜日ドロップダウンの選択で置換（「なし」のときは空文字で削除） */
                    pushSegment(oldParts.sep3, oldParts.sep3);
                    pushSegment(oldParts.suffix, explicitWeekdayText);
                } else {
                    /* 元 suffix なし：sep3（"日"）末尾に曜日を伸ばす */
                    pushSegment(oldParts.sep3, oldParts.sep3 + explicitWeekdayText);
                }
            } else {
                /* dot / slash / r-dot / r-slash：曜日サフィックスを含めて日付全体を1セグメントとして置換する */
                var preservedText = buildPreservedFormatText(matchInfo);
                if (preservedText !== null) {
                    segments = [{ offset: 0, oldLen: matchInfo.text.length, newText: preservedText }];
                }
            }
            return segments;
        }

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

        var operationCount = 0;
        var errorCount = 0;

        for (var frameIdxKey in matchesByFrame) {
            if (!matchesByFrame.hasOwnProperty(frameIdxKey)) continue;
            var targetIdx = parseInt(frameIdxKey, 10);
            var frameMatches = matchesByFrame[frameIdxKey];
            /* 後ろから置換すれば前方の matchIndex はずれない */
            frameMatches.sort(function (a, b) { return b.matchIndex - a.matchIndex; });

            try {
                for (var orderIdx = 0; orderIdx < frameMatches.length; orderIdx++) {
                    var currentMatch = frameMatches[orderIdx];
                    if (formatChoice === 'preserve' && preserveNumberFormat) {
                        var segments = buildReplacementSegments(currentMatch);
                        if (!segments) continue;
                        for (var segIdx = segments.length - 1; segIdx >= 0; segIdx--) {
                            var seg = segments[segIdx];
                            operationCount += replaceMatchPreserveStyle(
                                textFrames[targetIdx],
                                currentMatch.matchIndex + seg.offset,
                                seg.oldLen,
                                seg.newText
                            );
                        }
                    } else {
                        var fullText = buildReplacementText(currentMatch);
                        if (fullText === null) continue;
                        operationCount += replaceMatchPreserveStyle(
                            textFrames[targetIdx],
                            currentMatch.matchIndex,
                            currentMatch.text.length,
                            fullText
                        );
                    }
                }
            } catch (eReplace) {
                errorCount++;
            }
        }

        /*
           チェック済みマッチに紐づく曜日ペア（同一グループ内の曜日のみフレーム）を更新。
           曜日ドロップダウンが「なし」の場合は、空文字化せず連動フレームを更新しない。
           各ペアフレームは独立した TextFrame のままで、日付フレームには連結しない。
        */
        for (var pairMatchIndex = 0; pairMatchIndex < foundMatches.length; pairMatchIndex++) {
            var dateMatchWithPairs = foundMatches[pairMatchIndex];
            if (!dateMatchWithPairs.checkbox || !dateMatchWithPairs.checkbox.value) continue;
            if (!dateMatchWithPairs.weekdayPairs) continue;
            for (var weekdayPairIndex = 0; weekdayPairIndex < dateMatchWithPairs.weekdayPairs.length; weekdayPairIndex++) {
                var weekdayPair = dateMatchWithPairs.weekdayPairs[weekdayPairIndex];
                try {
                    /* 元のスタイルを維持しつつ、新しい曜日に置換。「なし」の場合は連動フレームを更新しない */
                    if (weekdayChoice === 'none') continue;
                    var newWeekdayText = buildExplicitWeekday(weekdayPair.style, newYear, newMonth, newDay);
                    if (newWeekdayText === "") continue;
                    var oldWeekdayText = String(weekdayPair.frame.contents);
                    if (oldWeekdayText === newWeekdayText) continue;
                    operationCount += replaceMatchPreserveStyle(weekdayPair.frame, 0, oldWeekdayText.length, newWeekdayText);
                } catch (eWeekdayPair) {
                    errorCount++;
                }
            }
        }

        return { operations: operationCount, errors: errorCount, selectedCount: selectedCount };
    }

    /* プレビューを巻き戻す（適用した文字操作数だけ app.undo() を呼ぶ） */
    function revertPreview() {
        if (!isPreviewApplied) return;
        for (var i = 0; i < previewOpCount; i++) {
            try { app.undo(); } catch (eUndo) { break; }
        }
        isPreviewApplied = false;
        previewOpCount = 0;
        try { app.redraw(); } catch (eRedraw) { }
    }

    /* プレビューを適用（呼び出し前に入力が有効であることを確認しておく） */
    function applyPreview() {
        var result = performReplacement();
        previewOpCount = result.operations;
        isPreviewApplied = result.operations > 0;
        try { app.redraw(); } catch (eRedraw) { }
    }

    /* 入力やチェック変更時の自動更新（プレビュー ON 時のみ巻き戻し→再適用） */
    function refreshPreview() {
        if (!previewCheckbox.value) return;
        revertPreview();
        if (!isInputValid()) return;
        applyPreview();
    }

    /* 「数字の書式を保持」は preserve 選択時のみ有効。フォーマット切替で enabled を連動させる */
    function updatePreserveNumberFormatEnabled() {
        var fc = getDropdownValue(formatDropdown, FORMAT_VALUES, 'preserve');
        preserveNumberFormatCheckbox.enabled = (fc === 'preserve');
    }
    updatePreserveNumberFormatEnabled();

    /* フォーマット選択 / 曜日 / 数字書式保持の変更でもプレビューを更新 */
    formatDropdown.onChange = function () {
        updatePreserveNumberFormatEnabled();
        refreshPreview();
    };
    weekdayDropdown.onChange = function () {
        updateCurrentWeekdayChoice();
        refreshPreview();
    };
    preserveNumberFormatCheckbox.onClick = function () { refreshPreview(); };

    /* プレビューチェックボックス本体 */
    previewCheckbox.onClick = function () {
        if (previewCheckbox.value) {
            var err = getValidationErrorMessage();
            if (err) {
                alert(err);
                previewCheckbox.value = false;
                return;
            }
            applyPreview();
        } else {
            revertPreview();
        }
    };

    /* OK / キャンセル ボタン */
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";

    var cancelButton = buttonGroup.add("button", undefined, "キャンセル");
    var okButton = buttonGroup.add("button", undefined, "OK");

    okButton.onClick = function () {
        var err = getValidationErrorMessage();
        if (err) {
            alert(err);
            return;
        }
        dialog.close(1);
    };

    cancelButton.onClick = function () {
        if (isPreviewApplied) revertPreview();
        dialog.close(0);
    };

    var dialogResult = dialog.show();

    if (dialogResult !== 1) {
        return;
    }

    if (foundMatches.length === 0) {
        alert("置換対象の日付がありません。");
        return;
    }

    // =========================================
    // 置換処理
    // =========================================

    /* OK時はプレビュー結果をそのまま確定せず、いったん戻してから本番置換を再実行する */
    if (isPreviewApplied) {
        revertPreview();
    }

    var result = performReplacement();

    if (result.selectedCount === 0) {
        alert("置換する項目が選択されていません。");
        return;
    }

    if (result.errors > 0) {
        alert(formatErrorMessage(result.errors));
    }
})();