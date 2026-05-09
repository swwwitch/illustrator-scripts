#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

概要

ドキュメント内のテキストフレームから日付を検索し、選択した項目のみを置換するスクリプト。

対応する日付形式：
  - YYYY年M月D日（曜日サフィックス付きも可：金曜日 / (金) / （金））
  - YYYY.M.D
  - YYYY/M/D
  - 令和Y.M.D
  - RY.M.D

置換時は元の表記スタイル（区切り文字、元号有無、和暦↔西暦の換算）を保持する。
検索結果はアートボードごとにグループ化して表示する。

*/

(function () {

    // =========================================
    // バージョン
    // =========================================

    var SCRIPT_VERSION = "v1.0.0";

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

    /* ↑↓キーで値を増減（Shiftで±10、Optionで±0.1） */
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
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
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
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                value = Math.round(value * 10) / 10;
            } else {
                value = Math.round(value);
            }

            editText.text = value;

            if (typeof onValueChanged === "function") {
                onValueChanged();
            }
        });
    }

    /* 一致した文字列の曜日サフィックスの形式を判定（JP形式のみ対象） */
    function detectWeekdaySuffixStyle(matchText) {
        if (/[日月火水木金土]曜日$/.test(matchText)) return 'long';   /* 例：金曜日 */
        if (/\([日月火水木金土]\)$/.test(matchText)) return 'half';   /* 例：(金) */
        if (/（[日月火水木金土]）$/.test(matchText)) return 'full';   /* 例：（金） */
        return null;
    }

    /* 令和元年 = 2019 年（令和Y = 西暦 Y + 2018） */
    var REIWA_BASE_YEAR = 2018;

    function reiwaToGregorian(reiwaYear) { return reiwaYear + REIWA_BASE_YEAR; }
    function gregorianToReiwa(gregorianYear) { return gregorianYear - REIWA_BASE_YEAR; }

    /*
       マッチ文字列を解析し、形式種別と各構成要素を返す。
       戻り値の主要フィールド：
         format: 'jp' | 'dot' | 'slash' | 'reiwa-dot' | 'r-dot'
         year, month, day: 西暦の数値（令和形式の場合は変換後）
         parts: 元テキストの分解（prefix, year, sep1, month, sep2, day, sep3, suffix）
    */
    function detectFormatAndParse(matchText) {
        var m;

        /* 令和Y.M.D */
        m = String(matchText).match(/^(令和)([0-9]{1,2})\.([0-9]{1,2})\.([0-9]{1,2})$/);
        if (m) {
            return {
                format: 'reiwa-dot', era: 'reiwa',
                year:  reiwaToGregorian(parseInt(m[2], 10)),
                month: parseInt(m[3], 10),
                day:   parseInt(m[4], 10),
                parts: { prefix: m[1], year: m[2], sep1: ".", month: m[3], sep2: ".", day: m[4], sep3: "", suffix: "" }
            };
        }

        /* RY.M.D */
        m = String(matchText).match(/^(R)([0-9]{1,2})\.([0-9]{1,2})\.([0-9]{1,2})$/);
        if (m) {
            return {
                format: 'r-dot', era: 'reiwa',
                year:  reiwaToGregorian(parseInt(m[2], 10)),
                month: parseInt(m[3], 10),
                day:   parseInt(m[4], 10),
                parts: { prefix: m[1], year: m[2], sep1: ".", month: m[3], sep2: ".", day: m[4], sep3: "", suffix: "" }
            };
        }

        /* YYYY年M月D日（曜日サフィックスは末尾に残す） */
        m = String(matchText).match(/^([0-9]{4})(年)([0-9]{1,2})(月)([0-9]{1,2})(日)(.*)$/);
        if (m) {
            return {
                format: 'jp', era: null,
                year:  parseInt(m[1], 10),
                month: parseInt(m[3], 10),
                day:   parseInt(m[5], 10),
                parts: { prefix: "", year: m[1], sep1: m[2], month: m[3], sep2: m[4], day: m[5], sep3: m[6], suffix: m[7] || "" }
            };
        }

        /* YYYY.M.D */
        m = String(matchText).match(/^([0-9]{4})\.([0-9]{1,2})\.([0-9]{1,2})$/);
        if (m) {
            return {
                format: 'dot', era: null,
                year:  parseInt(m[1], 10),
                month: parseInt(m[2], 10),
                day:   parseInt(m[3], 10),
                parts: { prefix: "", year: m[1], sep1: ".", month: m[2], sep2: ".", day: m[3], sep3: "", suffix: "" }
            };
        }

        /* YYYY/M/D */
        m = String(matchText).match(/^([0-9]{4})\/([0-9]{1,2})\/([0-9]{1,2})$/);
        if (m) {
            return {
                format: 'slash', era: null,
                year:  parseInt(m[1], 10),
                month: parseInt(m[2], 10),
                day:   parseInt(m[3], 10),
                parts: { prefix: "", year: m[1], sep1: "/", month: m[2], sep2: "/", day: m[3], sep3: "", suffix: "" }
            };
        }

        return null;
    }

    /* 日付からの曜日 1 文字 */
    function getJaWeekdayChar(year, month, day) {
        var checkDate = new Date(year, month - 1, day);
        if (isNaN(checkDate.getTime())) return "";
        var jaDays = ["日", "月", "火", "水", "木", "金", "土"];
        return jaDays[checkDate.getDay()];
    }

    /* 指定した形式で曜日サフィックス文字列を組み立て */
    function buildWeekdaySuffix(style, weekdayChar) {
        if (!style || !weekdayChar) return "";
        if (style === 'long') return weekdayChar + "曜日";
        if (style === 'half') return "(" + weekdayChar + ")";
        if (style === 'full') return "（" + weekdayChar + "）";
        return "";
    }

    /* テキストフレーム内の一致範囲を書式を保持したまま置換 */
    function replaceMatchPreserveStyle(textFrame, matchStart, oldLen, newText) {
        if (oldLen <= 0) return;
        var newLen = newText.length;
        var minLen = Math.min(oldLen, newLen);

        /* 重複部分は文字単位で書き換え。各 character.contents の代入は元の文字属性を維持 */
        for (var i = 0; i < minLen; i++) {
            textFrame.characters[matchStart + i].contents = newText.charAt(i);
        }

        if (newLen > oldLen) {
            /* 末尾文字に追加分の文字列を含めると、元の書式を引き継いだまま文字を増やせる */
            var lastIdx = matchStart + oldLen - 1;
            textFrame.characters[lastIdx].contents = newText.charAt(oldLen - 1) + newText.substring(oldLen);
        } else if (newLen < oldLen) {
            /* 余った古い文字を末尾から削除 */
            for (var ri = oldLen - 1; ri >= newLen; ri--) {
                textFrame.characters[matchStart + ri].remove();
            }
        }
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

    /* 曜日表示（例：（金）） */
    function getWeekdayLabel(year, month, day) {
        if (isNaN(year) || isNaN(month) || isNaN(day)) return "";
        var checkDate = new Date(year, month - 1, day);
        if (isNaN(checkDate.getTime())) return "";
        var jaDays = ["日", "月", "火", "水", "木", "金", "土"];
        return "（" + jaDays[checkDate.getDay()] + "）";
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
         2026年5月8日 / 2026年5月8日金曜日 / 2026年5月8日(金) / 2026年5月8日（金）
         2026.5.8 / 2026/5/8
         令和8.5.8 / R8.5.8
       元号系を先に列挙して優先マッチさせる。
    */
    var dateRegex = /令和[0-9]{1,2}\.[0-9]{1,2}\.[0-9]{1,2}|R[0-9]{1,2}\.[0-9]{1,2}\.[0-9]{1,2}|[0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日(?:[日月火水木金土]曜日|\([日月火水木金土]\)|（[日月火水木金土]）)?|[0-9]{4}\.[0-9]{1,2}\.[0-9]{1,2}|[0-9]{4}\/[0-9]{1,2}\/[0-9]{1,2}/g;

    var foundMatches = [];
    var textFrames = doc.textFrames;

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
            var parsed = detectFormatAndParse(match[0]);
            if (!parsed) continue;
            foundMatches.push({
                frameIndex: frameIdx,
                text: match[0],
                matchIndex: match.index,
                artboardIndex: artboardIdx,
                parsed: parsed,
                weekdaySuffixStyle: parsed.format === 'jp' ? detectWeekdaySuffixStyle(match[0]) : null,
                checkbox: null
            });
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

    /* チェックボックスに Option＋クリックで全切替を割り当てる */
    function bindOptionClickToggleAll(checkboxControl) {
        checkboxControl.addEventListener("click", function (event) {
            if (event.altKey) {
                var newValue = checkboxControl.value;
                for (var idx = 0; idx < allCheckboxes.length; idx++) {
                    allCheckboxes[idx].value = newValue;
                }
            }
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
                var checkbox = checkboxRow.add("checkbox", undefined, artboardItems[itemIdx].text);
                checkbox.value = true;
                /* チェックボックスのラベル切れ対策 */
                checkbox.preferredSize.width = 180;
                artboardItems[itemIdx].checkbox = checkbox;
                allCheckboxes.push(checkbox);
                bindOptionClickToggleAll(checkbox);
            }
        }
    }

    /* 置換後の日付入力パネル：［2026］年［5］月［8］日 形式 */
    var replacementPanel = dialog.add("panel", undefined, "置換する日付");
    setupPanel(replacementPanel, 6);

    var dateInputGroup = replacementPanel.add("group");
    dateInputGroup.orientation = "row";

    var yearInput = dateInputGroup.add("edittext", undefined, String(defaultYear));
    yearInput.characters = 5;
    dateInputGroup.add("statictext", undefined, "年");

    var monthInput = dateInputGroup.add("edittext", undefined, String(defaultMonth));
    monthInput.characters = 3;
    dateInputGroup.add("statictext", undefined, "月");

    var dayInput = dateInputGroup.add("edittext", undefined, String(defaultDay));
    dayInput.characters = 3;
    dateInputGroup.add("statictext", undefined, "日");

    /*
       出力フォーマット選択。
       「元の形式を保持」を選ぶと、各マッチの元形式と元の曜日付き／なしをそのまま維持して置換する。
       それ以外は、選択した形式で全マッチを統一して書き換える。
    */
    var FORMAT_VALUES = ["preserve", "jp", "jp-full", "jp-long", "dot", "slash", "reiwa-dot", "r-dot"];
    var FORMAT_LABELS = [
        "元の形式を保持",
        "YYYY年M月D日",
        "YYYY年M月D日（金）",
        "YYYY年M月D日金曜日",
        "YYYY.M.D",
        "YYYY/M/D",
        "令和Y.M.D",
        "RY.M.D"
    ];
    var formatRow = replacementPanel.add("group");
    formatRow.orientation = "row";
    formatRow.add("statictext", undefined, "フォーマット：");
    var formatDropdown = formatRow.add("dropdownlist", undefined, FORMAT_LABELS);
    formatDropdown.selection = 0;

    /* 数字（年・月・日）の文字書式を保持するか。OFF にするとマッチ範囲を一括置換し、
       新しい文字はマッチ先頭の書式に統一される */
    var preserveNumberFormatCheckbox = replacementPanel.add("checkbox", undefined, "数字の書式を保持する");
    preserveNumberFormatCheckbox.value = true;

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

    /* 入力値からチェック用の表示を更新 */
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
    }

    yearInput.onChange = refreshCheckLabels;
    monthInput.onChange = refreshCheckLabels;
    dayInput.onChange = refreshCheckLabels;

    changeValueByArrowKey(yearInput, refreshCheckLabels);
    changeValueByArrowKey(monthInput, refreshCheckLabels);
    changeValueByArrowKey(dayInput, refreshCheckLabels);

    refreshCheckLabels();

    /* OK / キャンセル ボタン */
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";

    var cancelButton = buttonGroup.add("button", undefined, "キャンセル");
    var okButton = buttonGroup.add("button", undefined, "OK");

    okButton.onClick = function () {
        var inputYear = parseInt(yearInput.text, 10);
        var inputMonth = parseInt(monthInput.text, 10);
        var inputDay = parseInt(dayInput.text, 10);

        if (isNaN(inputYear) || isNaN(inputMonth) || isNaN(inputDay)) {
            alert("年・月・日は半角数字で入力してください。");
            return;
        }

        if (inputMonth < 1 || inputMonth > 12) {
            alert("月は1〜12で入力してください。");
            return;
        }

        if (inputDay < 1 || inputDay > 31) {
            alert("日は1〜31で入力してください。");
            return;
        }

        /* 実在する日付かチェック */
        var checkDate = new Date(inputYear, inputMonth - 1, inputDay);
        if (
            checkDate.getFullYear() !== inputYear ||
            checkDate.getMonth() !== inputMonth - 1 ||
            checkDate.getDate() !== inputDay
        ) {
            alert("存在しない日付です。");
            return;
        }

        dialog.close(1);
    };

    cancelButton.onClick = function () {
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

    var newYear = parseInt(yearInput.text, 10);
    var newMonth = parseInt(monthInput.text, 10);
    var newDay = parseInt(dayInput.text, 10);

    var newWeekdayChar = getJaWeekdayChar(newYear, newMonth, newDay);
    var formatChoice = FORMAT_VALUES[formatDropdown.selection.index];
    var preserveNumberFormat = preserveNumberFormatCheckbox.value;

    /* マッチの元形式を維持した置換テキストを構築（元の曜日サフィックスはそのまま再現） */
    function buildPreservedFormatText(matchInfo) {
        var parsed = matchInfo.parsed;
        if (!parsed) return null;
        var oldParts = parsed.parts;

        var newYearText = (parsed.era === 'reiwa') ? String(gregorianToReiwa(newYear)) : String(newYear);

        var text =
            oldParts.prefix +
            newYearText +
            oldParts.sep1 +
            String(newMonth) +
            oldParts.sep2 +
            String(newDay) +
            oldParts.sep3;

        if (parsed.format === 'jp' && oldParts.suffix.length > 0) {
            text += buildWeekdaySuffix(matchInfo.weekdaySuffixStyle, newWeekdayChar);
        }
        return text;
    }

    /* 指定フォーマットでの置換テキストを構築（全マッチで形式を統一） */
    function buildExplicitFormatText(format) {
        var reiwaYear = gregorianToReiwa(newYear);
        switch (format) {
            case 'jp':       return newYear + "年" + newMonth + "月" + newDay + "日";
            case 'jp-full':  return newYear + "年" + newMonth + "月" + newDay + "日（" + newWeekdayChar + "）";
            case 'jp-long':  return newYear + "年" + newMonth + "月" + newDay + "日" + newWeekdayChar + "曜日";
            case 'dot':      return newYear + "." + newMonth + "." + newDay;
            case 'slash':    return newYear + "/" + newMonth + "/" + newDay;
            case 'reiwa-dot': return "令和" + reiwaYear + "." + newMonth + "." + newDay;
            case 'r-dot':    return "R" + reiwaYear + "." + newMonth + "." + newDay;
        }
        return null;
    }

    /* 最終的な置換テキスト：「元の形式を保持」なら原形に合わせ、それ以外は選択フォーマット */
    function buildReplacementText(matchInfo) {
        if (formatChoice === 'preserve') {
            return buildPreservedFormatText(matchInfo);
        }
        return buildExplicitFormatText(formatChoice);
    }

    /*
       マッチを構成要素ごとに分解し、各セグメント単位の置換情報を返す。
       こうすることで「5」→「12」のように桁数が変化しても、各部分は元の文字属性を保持できる。
       元号付きの場合は西暦↔令和の換算も行う。
       戻り値は { offset, oldLen, newText } の配列で、offset はマッチ先頭からの相対位置。
       セグメント方式は「元の形式を保持」かつ「数字の書式を保持する」が両方 ON の場合のみ使用。
    */
    function buildReplacementSegments(matchInfo) {
        var parsed = matchInfo.parsed;
        if (!parsed) return null;

        var oldParts = parsed.parts;
        var segments = [];
        var pos = 0;

        function pushSegment(oldText, newText) {
            if (oldText.length === 0) {
                /* 0 文字セグメントには書式付き挿入ができないので、呼び出し側で前のセグメントに連結する想定 */
                return;
            }
            segments.push({ offset: pos, oldLen: oldText.length, newText: newText });
            pos += oldText.length;
        }

        var newYearText = (parsed.era === 'reiwa') ? String(gregorianToReiwa(newYear)) : String(newYear);

        pushSegment(oldParts.prefix, oldParts.prefix);   /* "令和" / "R" / "" */
        pushSegment(oldParts.year, newYearText);
        pushSegment(oldParts.sep1, oldParts.sep1);       /* "年" / "." / "/" */
        pushSegment(oldParts.month, String(newMonth));
        pushSegment(oldParts.sep2, oldParts.sep2);
        pushSegment(oldParts.day, String(newDay));

        if (parsed.format === 'jp') {
            if (oldParts.suffix.length > 0) {
                /* 既存の曜日サフィックスは同じ形式で更新 */
                pushSegment(oldParts.sep3, oldParts.sep3);
                pushSegment(oldParts.suffix, buildWeekdaySuffix(matchInfo.weekdaySuffixStyle, newWeekdayChar));
            } else {
                pushSegment(oldParts.sep3, oldParts.sep3);
            }
        }

        return segments;
    }

    /* チェック済みの結果のみテキストフレーム単位でまとめる */
    var matchesByFrame = {};
    var selectedCount = 0;
    for (var resultIdx = 0; resultIdx < foundMatches.length; resultIdx++) {
        if (!foundMatches[resultIdx].checkbox || !foundMatches[resultIdx].checkbox.value) continue;
        var targetFrameIdx = foundMatches[resultIdx].frameIndex;
        if (!matchesByFrame[targetFrameIdx]) matchesByFrame[targetFrameIdx] = [];
        matchesByFrame[targetFrameIdx].push(foundMatches[resultIdx]);
        selectedCount++;
    }

    if (selectedCount === 0) {
        alert("置換する項目が選択されていません。");
        return;
    }

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
                /*
                   「元の形式を保持」かつ「数字の書式を保持する」が両方 ON のときだけ
                   セグメント単位の置換で各部分の文字属性を細かく保ったまま書き換える。
                   それ以外は、マッチ範囲全体を 1 回で置換する（書式は先頭文字に揃う）。
                */
                if (formatChoice === 'preserve' && preserveNumberFormat) {
                    var segments = buildReplacementSegments(currentMatch);
                    if (!segments) continue;
                    for (var segIdx = segments.length - 1; segIdx >= 0; segIdx--) {
                        var seg = segments[segIdx];
                        replaceMatchPreserveStyle(
                            textFrames[targetIdx],
                            currentMatch.matchIndex + seg.offset,
                            seg.oldLen,
                            seg.newText
                        );
                    }
                } else {
                    var fullText = buildReplacementText(currentMatch);
                    if (fullText === null) continue;
                    replaceMatchPreserveStyle(
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

    /* エラー時のみアラート表示（成功時は通知しない） */
    if (errorCount > 0) {
        alert(formatErrorMessage(errorCount));
    }
})();
