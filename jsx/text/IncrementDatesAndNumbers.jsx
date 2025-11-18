#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

var SCRIPT_VERSION = "v1.2";

/*
 * 値の増減スクリプト / Increment or decrement dates and numbers
 *
 * 選択中のテキスト内に含まれる日付・曜日・連番・数値などを一括して増減します。
 * 増減値は整数のみを許容し、小数は使用できません。
 *
 * 対象となる主な文字列パターン：
 * - 日付（和文）：2025年11月21日、2025年11月21日（金）
 * - 日付（スラッシュ区切り）：2025/11/21、2025/11/21（金）
 * - 日付（ドット区切り）：2025.11.21、29.3.2、29.3
 * - 元号付き日付：令和7年3月21日（金）など
 * - 時刻：19:00 など
 * - 曜日単体：日・月・火・水・木・金・土、Sun〜Sat、㊐〜㊏
 * - 西暦（年のみ）：2025 など
 * - 一般的な数値：123、1,234、100.5 など
 *
 * This script increments or decrements dates, weekdays, sequence numbers, and numeric values
 * inside selected text objects. The step value accepts integers only; decimals are not allowed.
 *
 * Typical target string patterns include:
 * - Japanese dates: 2025年11月21日, 2025年11月21日（金）
 * - Slash dates: 2025/11/21, 2025/11/21(Fri)
 * - Dot-separated dates: 2025.11.21, 29.3.2, 29.3
 * - Era-based dates: 令和7年3月21日（金）, etc.
 * - Times: 19:00, etc.
 * - Standalone weekdays: 日〜土, Sun–Sat, circled symbols ㊐〜㊏
 * - Standalone years: 2025
 * - Generic numeric values: 123, 1,234, 100.5
 *
 * 更新日 / Last update: 2025-11-18
 */

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var STEP_VALUE = 1;
var SHIFT_MODE = "day";

// ドット区切り2要素（例：12.1）の解釈モード
// Mode for interpreting 2-part dot patterns like 12.1 ("number" or "date")
var DOT2_MODE = "number";

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

var LABELS = {
    dialogTitle: {
        ja: "値の増減",
        en: "Increment / Decrement values"
    },
    originalLabel: {
        ja: "オリジナル：",
        en: "Original:"
    },
    resultLabel: {
        ja: "結果：",
        en: "Result:"
    },
    dirPanelTitle: {
        ja: "増減",
        en: "Increment"
    },
    stepLabel: {
        ja: "値",
        en: "Value"
    },
    targetLabel: {
        ja: "対象",
        en: "Target"
    },
    modePanelTitle: {
        ja: "種別",
        en: "Type"
    },
    modeNumber: {
        ja: "数字",
        en: "Number"
    },
    modeDate: {
        ja: "日付",
        en: "Date"
    },
    btnOK: {
        ja: "OK",
        en: "OK"
    },
    btnCancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    errorTitle: {
        ja: "エラー",
        en: "Error"
    },
    errorGeneric: {
        ja: "エラーが発生しました：",
        en: "An error occurred:"
    }
};

function L(key) {
    if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
    if (LABELS[key] && LABELS[key].en) return LABELS[key].en;
    return key;
}

function findFirstTextFrame(item) {
    if (!item) return null;
    if (item.typename === "TextFrame") return item;
    if (item.typename === "GroupItem") {
        var children = item.pageItems;
        for (var i = 0; i < children.length; i++) {
            var found = findFirstTextFrame(children[i]);
            if (found) return found;
        }
    }
    return null;
}

function findFirstTextFrameInSelection(sel) {
    if (!sel || sel.length === 0) return null;
    for (var i = 0; i < sel.length; i++) {
        var found = findFirstTextFrame(sel[i]);
        if (found) return found;
    }
    return null;
}

function showDialog() {
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];
    dlg.margins = 15;
    dlg.spacing = 10;

    var hasYMDTarget = false;
    var hasDot2Ambiguous = false;
    var labelYear = "";
    var labelMonth = "";
    var labelDay = "";
    var originalSample = "";

    try {
        var doc = app.activeDocument;
        var sel = doc.selection;
        if (sel && sel.length > 0) {
            var target = findFirstTextFrameInSelection(sel);
            if (target && target.typename === "TextFrame") {
                var txtContent = target.contents.replace(/^\s+|\s+$/g, "");
                var lineBreak = String.fromCharCode(13);
                originalSample = txtContent.split(lineBreak)[0];

                var mYMD = txtContent.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日(?:[（(［\[]?(?:日|月|火|水|木|金|土)[）)\]］]?)?$/);
                if (mYMD) {
                    // 2025年11月21日のような和文年月日
                    hasYMDTarget = true;
                    labelYear = mYMD[1];
                    labelMonth = mYMD[2];
                    labelDay = mYMD[3];
                } else {
                    // 2025/11/21 や 2025/11/21（金）のようなスラッシュ区切り
                    // Slash-separated date with optional weekday suffix
                    var mSlash = txtContent.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})(?:[ 　\t]*[（(［\[]?(?:日|月|火|水|木|金|土|Sun|Mon|Tue|Wed|Thu|Fri|Sat)[）)\]］]?)?$/);
                    if (mSlash) {
                        hasYMDTarget = true;
                        labelYear = mSlash[1];
                        labelMonth = mSlash[2];
                        labelDay = mSlash[3];
                    } else {
                        // 29.3.2 / 29.4 などのドット区切り
                        var mDot3 = txtContent.match(/^(\d+)\.(\d+)\.(\d+)$/);
                        var mDot2 = txtContent.match(/^(\d+)\.(\d+)$/);
                        if (mDot3) {
                            hasYMDTarget = true;
                            labelYear = mDot3[1];
                            labelMonth = mDot3[2];
                            labelDay = mDot3[3];
                        } else if (mDot2) {
                            hasYMDTarget = true;
                            labelYear = mDot2[1];
                            labelMonth = mDot2[2];
                            labelDay = "";
                            // 2要素のドット区切り（例：12.1）は「数字／日付」のどちらにもなり得るため
                            // モード選択パネルを表示する
                            // Two-part dot patterns (e.g. 12.1) can be numeric or date, so enable mode selection
                            hasDot2Ambiguous = true;
                        } else {
                            // 時刻 19:00 のようなパターンも分割対象とする（年=時、月=分として扱う）
                            // Treat time like 19:00 as a split target (map year->hour, month->minute)
                            var mTime = txtContent.match(/^(\d{1,2}):(\d{2})$/);
                            if (mTime) {
                                hasYMDTarget = true;
                                labelYear = mTime[1]; // hour
                                labelMonth = mTime[2]; // minute
                                labelDay = "";
                            }
                        }
                    }
                }
            }
        }
    } catch (e) {}


    var grpPreview = dlg.add('group');
    grpPreview.orientation = 'column';
    grpPreview.alignChildren = ['left', 'top'];
    grpPreview.margins = [13, 10, 0, 10];
    grpPreview.spacing = 6;

    var grpOriginal = grpPreview.add('group');
    grpOriginal.orientation = 'row';
    grpOriginal.alignChildren = ['left', 'center'];
    grpOriginal.spacing = 10;

    var lblOriginal = grpOriginal.add('statictext', undefined, L('originalLabel'));
    var stOriginalValue = grpOriginal.add('statictext', undefined, originalSample || "");

    var grpResult = grpPreview.add('group');
    grpResult.orientation = 'row';
    grpResult.alignChildren = ['left', 'center'];
    grpResult.spacing = 5;

    var lblResult = grpResult.add('statictext', undefined, L('resultLabel'));
    var stResultValue = grpResult.add('statictext', undefined, "");
    var stCalcValue = grpResult.add('statictext', undefined, "");

    // ラベル幅と計算式領域
    try {
        var g = dlg.graphics;

        // 「オリジナル」「結果」ラベルの幅を揃えて右揃えに
        // Measure label widths and align them right with the same width
        var wOrig = g.measureString(lblOriginal.text)[0];
        var wRes = g.measureString(lblResult.text)[0];
        var maxW = Math.ceil(Math.max(wOrig, wRes));

        lblOriginal.preferredSize = [maxW, -1];
        lblResult.preferredSize = [maxW, -1];
        lblOriginal.justify = 'right';
        lblResult.justify = 'right';

        // ラベルと値の間のスペースを一定にする
        grpOriginal.spacing = 6;
        grpResult.spacing = 6;

        // 結果値用に十分な幅を確保（オリジナルが短い場合は日付相当の長さを基準にする）
        // Reserve enough width for the result; if the original is short (e.g. "12.1"),
        // fall back to a typical date length like "0000年00月00日" to avoid truncation.
        var sampleResult = originalSample || "0000年00月00日";
        if (sampleResult.length < 8) {
            // 「12.1」など短いケースは、日付想定のサンプルで幅を確保
            sampleResult = "0000年00月00日";
        }
        var wResult = g.measureString(sampleResult)[0];
        stResultValue.preferredSize = [Math.ceil(wResult), -1];

        // 右側の数値表示用の幅は最小限にしてダイアログ全体の幅を抑える
        // Keep the right-side numeric field narrow so the dialog isn't too wide
        // stCalcValue.preferredSize = [30, -1];
    } catch (e) {}

    function updateResultPreview() {
        if (!originalSample || originalSample === "") {
            stResultValue.text = "";
            if (stCalcValue) stCalcValue.text = "";
            return;
        }
        try {
            // 現在の設定で1行分を変換 / Transform one line with current settings
            var previewText = updateLine(originalSample);

            // 変換後の全文をそのまま結果として表示
            // Always show the full transformed text (e.g. 2025年12月18日)
            var resultDisplay = previewText;

            // 結果ラベルに表示 / Show in the Result field
            stResultValue.text = resultDisplay;

            // 年月日ターゲットの場合は右側の計算結果は使わない / no extra numeric on the right for Y/M/D target
            if (hasYMDTarget) {
                if (stCalcValue) stCalcValue.text = "";
                return;
            }

            // それ以外のケースでは、右側に数値結果を表示 / For other cases, show numeric result on the right
            if (stCalcValue) {
                // オリジナルも結果も「純粋な数値のみ」の場合は右側には何も表示しない
                // If both original and result are pure numeric strings, don't show a duplicate value on the right
                var pureNumReg = /^\s*-?\d+(?:,\d{3})*(?:\.\d+)?\s*$/;
                if (pureNumReg.test(originalSample) && pureNumReg.test(previewText)) {
                    stCalcValue.text = "";
                    return;
                }

                var numPattern = /-?\d+(?:,\d{3})*(?:\.\d+)?/;
                var origMatch = originalSample.match(numPattern);
                var resMatch = previewText.match(numPattern);

                if (origMatch && resMatch) {
                    var origRaw = origMatch[0].replace(/,/g, "").replace(/^\s+|\s+$/g, "");
                    var resRaw = resMatch[0].replace(/,/g, "").replace(/^\s+|\s+$/g, "");

                    var origNum = Number(origRaw);
                    var resNum = Number(resRaw);

                    if (!isNaN(origNum) && !isNaN(resNum)) {
                        // 計算結果のみを表示 / Show only the numeric result
                        stCalcValue.text = resNum;
                    } else {
                        stCalcValue.text = "";
                    }
                } else {
                    stCalcValue.text = "";
                }
            }
        } catch (e) {
            stResultValue.text = "";
            if (stCalcValue) stCalcValue.text = "";
        }
    }

    dlg.onShow = function() {
        try {
            edtStep.active = true;
        } catch (e) {}
        updateResultPreview();
    };

    var pnlDir = dlg.add('panel', undefined, L('dirPanelTitle'));
    pnlDir.orientation = 'column';
    pnlDir.alignChildren = ['left', 'top'];
    pnlDir.margins = [15, 20, 15, 13];
    pnlDir.spacing = 10;

    var grpStep = pnlDir.add('group');
    grpStep.orientation = 'row';
    grpStep.alignChildren = ['left', 'center'];
    grpStep.spacing = 5;
    grpStep.add('statictext', undefined, L('stepLabel'));
    var edtStep = grpStep.add('edittext', undefined, "1");
    edtStep.characters = 5;
    changeValueByArrowKey(edtStep);

    edtStep.onChanging = function() {
        var v = parseInt(edtStep.text, 10);
        if (isNaN(v)) return;
        STEP_VALUE = v;
        edtStep.text = String(v);
        updateResultPreview();
    };
    edtStep.onChange = edtStep.onChanging;

    // 「12.1」のようなドット2要素のとき、数字／日付の解釈を選択するパネル
    // When the pattern is like "12.1", let the user choose between numeric and date interpretation
    var rbModeNumber, rbModeDate;
    if (hasDot2Ambiguous) {
        var pnlMode = dlg.add('panel', undefined, L('modePanelTitle'));
        pnlMode.orientation = 'row';
        pnlMode.alignChildren = ['left', 'center'];
        pnlMode.margins = [15, 20, 15, 10];
        pnlMode.spacing = 15;

        rbModeNumber = pnlMode.add('radiobutton', undefined, L('modeNumber'));
        rbModeDate = pnlMode.add('radiobutton', undefined, L('modeDate'));

        // デフォルトは「数字」
        DOT2_MODE = "number";
        rbModeNumber.value = true;

        rbModeNumber.onClick = function() {
            DOT2_MODE = "number";
            updateResultPreview();
        };
        rbModeDate.onClick = function() {
            DOT2_MODE = "date";
            updateResultPreview();
        };
    }

    var rbTargetYear, rbTargetMonth, rbTargetDay;
    if (hasYMDTarget) {
        var pnlTarget = dlg.add('panel', undefined, L('targetLabel'));
        pnlTarget.orientation = 'column';
        pnlTarget.alignChildren = ['left', 'top'];
        pnlTarget.margins = [15, 20, 15, 10];
        pnlTarget.spacing = 2;

        if (labelYear !== "") rbTargetYear = pnlTarget.add('radiobutton', undefined, labelYear);
        if (labelMonth !== "") rbTargetMonth = pnlTarget.add('radiobutton', undefined, labelMonth);
        if (labelDay !== "") rbTargetDay = pnlTarget.add('radiobutton', undefined, labelDay);

        if (rbTargetDay) rbTargetDay.value = true;
        else if (rbTargetMonth) rbTargetMonth.value = true;
        else if (rbTargetYear) rbTargetYear.value = true;

        // ラジオボタンの選択に応じて SHIFT_MODE を更新し、必要なら増減値をリセットしてプレビューに反映
        // Update SHIFT_MODE according to the selected radio button, optionally reset step to 1, and refresh preview
        function applyTargetModeFromRadio(resetStep) {
            var mode = "day";
            if (rbTargetYear && rbTargetYear.value) {
                mode = "year";
            } else if (rbTargetMonth && rbTargetMonth.value) {
                mode = "month";
            } else if (rbTargetDay && rbTargetDay.value) {
                mode = "day";
            }
            SHIFT_MODE = mode;

            // 分割処理の対象を切り替えたときは、増減値を 1 にリセットし、編集状態にする
            // When switching the split target, reset step value to 1 and focus the step field
            if (resetStep && edtStep) {
                STEP_VALUE = 1;
                edtStep.text = "1";
                try {
                    edtStep.active = true;
                } catch (e) {}
            }

            updateResultPreview();
        }

        if (rbTargetYear) {
            rbTargetYear.onClick = function() {
                applyTargetModeFromRadio(true);
            };
        }
        if (rbTargetMonth) {
            rbTargetMonth.onClick = function() {
                applyTargetModeFromRadio(true);
            };
        }
        if (rbTargetDay) {
            rbTargetDay.onClick = function() {
                applyTargetModeFromRadio(true);
            };
        }

        // デフォルト選択の状態を SHIFT_MODE に反映（このタイミングでは増減値はリセットしない）
        // Sync initial selection to SHIFT_MODE without resetting the step value
        applyTargetModeFromRadio(false);
    }

    var grpButtons = dlg.add('group');
    grpButtons.orientation = 'row';
    grpButtons.alignChildren = ['right', 'center'];
    grpButtons.alignment = ['fill', 'bottom'];
    grpButtons.spacing = 10;

    var btnCancel = grpButtons.add('button', undefined, L('btnCancel'));
    var btnOK = grpButtons.add('button', undefined, L('btnOK'));

    btnOK.onClick = function() {
        var v = parseInt(edtStep.text, 10);
        if (isNaN(v)) v = 1;
        STEP_VALUE = v;
        edtStep.text = String(v);

        // ラジオボタンの状態を最終的な SHIFT_MODE に反映（増減値は維持）
        if (typeof applyTargetModeFromRadio === 'function') {
            applyTargetModeFromRadio(false);
        }

        updateResultPreview();
        main();
        dlg.close(1);
    };

    btnCancel.onClick = function() {
        dlg.close(0);
    };

    // レイアウト後にダイアログのレイアウトを確定 / Finalize layout before showing dialog
    dlg.layout.layout(true);

    dlg.center();
    dlg.show();
}

function main() {
    try {
        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        for (var i = 0; i < selectedItems.length; i++) {
            processSelectedItem(selectedItems[i]);
        }
    } catch (e) {
        alert(L("errorGeneric") + String.fromCharCode(13) + e.message, L("errorTitle"));
    }
}

function processSelectedItem(item) {
    if (!item) return;
    if (item.typename === "TextFrame") {
        var lineBreak = String.fromCharCode(13);
        var lines = item.contents.split(lineBreak);
        var updatedLines = [];
        for (var i = 0; i < lines.length; i++) {
            updatedLines.push(updateLine(lines[i]));
        }
        item.contents = updatedLines.join(lineBreak);
        return;
    }
    if (item.typename === "GroupItem") {
        var children = item.pageItems;
        for (var j = 0; j < children.length; j++) {
            processSelectedItem(children[j]);
        }
    }
}

function shiftDateByMode(date, stepInt) {
    if (SHIFT_MODE === "year") {
        date.setFullYear(date.getFullYear() + stepInt);
    } else if (SHIFT_MODE === "month") {
        date.setMonth(date.getMonth() + stepInt);
    } else {
        date.setDate(date.getDate() + stepInt);
    }
}

function updateLine(text) {
    try {
        var updatedText = text.replace(/^\s+|\s+$/g, "");
        var stepVal = (typeof STEP_VALUE === 'number' && !isNaN(STEP_VALUE)) ? STEP_VALUE : 1;
        var stepInt;

        // 0 のときは変化させない（stepInt を 0 に）/ When step is 0, make no change
        if (stepVal === 0) {
            stepInt = 0;
        } else if (stepVal > 0) {
            stepInt = Math.max(1, Math.round(stepVal));
        } else {
            stepInt = Math.min(-1, Math.round(stepVal));
        }

        var weekdayShift = stepInt % 7;
        if (weekdayShift < 0) weekdayShift += 7;

        var match = updatedText.match(/(\d{1,2})月(\d{1,2})日(㊐|㊊|㊋|㊌|㊍|㊎|㊏)/);
        if (match) {
            var m = parseInt(match[1], 10);
            var d = parseInt(match[2], 10);
            var date = new Date(CURRENT_YEAR, m - 1, d);
            shiftDateByMode(date, stepInt);
            var newSymbol = DAY_SYMBOLS[date.getDay()];
            var newStr = (date.getMonth() + 1) + "月" + date.getDate() + "日" + newSymbol;
            updatedText = updatedText.replace(match[0], newStr);
            return updatedText;
        }

        match = updatedText.match(/(\d{1,2})月(\d{1,2})日([（([\[]?)(日|月|火|水|木|金|土)([）)\]]?)/);
        if (match) {
            var date2 = new Date(CURRENT_YEAR, parseInt(match[1], 10) - 1, parseInt(match[2], 10));
            shiftDateByMode(date2, stepInt);
            var prefix = match[3] || "";
            var suffix = match[5] || "";
            var newStr2 = (date2.getMonth() + 1) + "月" + date2.getDate() + "日" + prefix + DAY_NAMES[date2.getDay()] + suffix;
            updatedText = updatedText.replace(match[0], newStr2);
            return updatedText;
        }

        match = updatedText.match(/(明治|大正|昭和|平成|令和)(\d{1,2})年(\d{1,2})月(\d{1,2})日([（([\[]?)(日|月|火|水|木|金|土)([）)\]]?)/);
        if (match) {
            var baseYear = ERA_MAP[match[1]];
            var date3 = new Date(baseYear + parseInt(match[2], 10) - 1, parseInt(match[3], 10) - 1, parseInt(match[4], 10));
            shiftDateByMode(date3, stepInt);
            var newEra = match[1];
            for (var k in ERA_MAP) {
                if (date3.getFullYear() >= ERA_MAP[k]) newEra = k;
            }
            var eraYear = date3.getFullYear() - ERA_MAP[newEra] + 1;
            var prefix3 = match[5] || "";
            var suffix3 = match[7] || "";
            var newStr3 = newEra + eraYear + "年" + (date3.getMonth() + 1) + "月" + date3.getDate() + "日" + prefix3 + DAY_NAMES[date3.getDay()] + suffix3;
            updatedText = updatedText.replace(match[0], newStr3);
            return updatedText;
        }

        match = updatedText.match(/(\d{4})([\/.])(\d{1,2})\2(\d{1,2})(\s*)([（([\[]?)(日|月|火|水|木|金|土|Sun|Mon|Tue|Wed|Thu|Fri|Sat)([）)\]]?)/);
        if (match) {
            var y = parseInt(match[1], 10);
            var sep = match[2];
            var m4 = parseInt(match[3], 10);
            var d4 = parseInt(match[4], 10);
            var space = match[5] || "";
            var prefix4 = match[6] || "";
            var dayStr = match[7];
            var suffix4 = match[8] || "";
            var date4 = new Date(y, m4 - 1, d4);
            shiftDateByMode(date4, stepInt);
            var newDateStr = date4.getFullYear() + sep + (date4.getMonth() + 1) + sep + date4.getDate();
            var newWeekday = /^[日月火水木金土]$/.test(dayStr) ? DAY_NAMES[date4.getDay()] : ENGLISH_DAYS[date4.getDay()];
            var originalStr = match[1] + sep + match[3] + sep + match[4] + space + prefix4 + dayStr + suffix4;
            var replacedStr = newDateStr + space + prefix4 + newWeekday + suffix4;
            updatedText = updatedText.replace(originalStr, replacedStr);
            return updatedText;
        }

        var fullDateRegexes = [{
                regex: /(\d{4})(\d{2})(\d{2})/,
                format: "plain"
            },
            {
                regex: /(\d{4})\/(\d{1,2})\/(\d{1,2})/,
                format: "slash"
            },
            {
                regex: /(\d{4})\.(\d{1,2})\.(\d{1,2})/,
                format: "dot"
            }
        ];
        for (var i = 0; i < fullDateRegexes.length; i++) {
            var r = fullDateRegexes[i];
            var m5 = updatedText.match(r.regex);
            if (m5) {
                var date5 = new Date(parseInt(m5[1], 10), parseInt(m5[2], 10) - 1, parseInt(m5[3], 10));
                shiftDateByMode(date5, stepInt);
                var replaced = (r.format === "plain") ?
                    pad(date5.getFullYear(), 4) + pad(date5.getMonth() + 1, 2) + pad(date5.getDate(), 2) :
                    date5.getFullYear() + (r.format === "slash" ? "/" : ".") + (date5.getMonth() + 1) + (r.format === "slash" ? "/" : ".") + date5.getDate();
                updatedText = updatedText.replace(r.regex, replaced);
                return updatedText;
            }
        }

        // 「2025年11月20日(木)」のように曜日付きの年月日
        // Full Japanese date with weekday in parentheses
        match = updatedText.match(/(\d{4})年(\d{1,2})月(\d{1,2})日([（(［\[]?)(日|月|火|水|木|金|土)([）)\]]?)/);
        if (match) {
            var y3 = parseInt(match[1], 10);
            var m7 = parseInt(match[2], 10);
            var d7 = parseInt(match[3], 10);
            var prefix5 = match[4] || ""; // 開きかっこ / opening bracket
            var suffix5 = match[6] || ""; // 閉じかっこ / closing bracket
            var date7 = new Date(y3, m7 - 1, d7);
            // SHIFT_MODE に応じて日・月・年を増減 / shift year, month, or day
            shiftDateByMode(date7, stepInt);
            var newStr5 = date7.getFullYear() + "年" + (date7.getMonth() + 1) + "月" + date7.getDate() + "日" + prefix5 + DAY_NAMES[date7.getDay()] + suffix5;
            updatedText = updatedText.replace(match[0], newStr5);
            return updatedText;
        }

        // 「2025年11月20日」のように曜日なしの年月日（直後にかっこが続かないものだけを対象）
        // Plain Japanese date without weekday; ensure not immediately followed by a bracket (to avoid double-handling cases like "2025年11月20日(木)")
        match = updatedText.match(/(\d{4})年(\d{1,2})月(\d{1,2})日(?![（(［\[])/);
        if (match) {
            var y2 = parseInt(match[1], 10);
            var m6 = parseInt(match[2], 10);
            var d6 = parseInt(match[3], 10);
            var date6 = new Date(y2, m6 - 1, d6);
            shiftDateByMode(date6, stepInt);
            var newStr4 = date6.getFullYear() + "年" + (date6.getMonth() + 1) + "月" + date6.getDate() + "日";
            updatedText = updatedText.replace(match[0], newStr4);
            return updatedText;
        }

        var symbolMatch = updatedText.match(/(㊐|㊊|㊋|㊌|㊍|㊎|㊏)/);
        if (symbolMatch) {
            var idx = DAY_SYMBOLS.indexOf(symbolMatch[1]);
            if (idx !== -1) {
                var newIdx = (idx + weekdayShift) % 7;
                updatedText = updatedText.replace(symbolMatch[1], DAY_SYMBOLS[newIdx]);
            }
        }

        var match2 = updatedText.match(/([（([\[]?)(日|月|火|水|木|金|土)([）)\]]?)/);
        if (match2) {
            var idx2 = DAY_NAMES.indexOf(match2[2]);
            if (idx2 !== -1) {
                var newIdx2 = (idx2 + weekdayShift) % 7;
                updatedText = updatedText.replace(match2[0], (match2[1] || "") + DAY_NAMES[newIdx2] + (match2[3] || ""));
            }
        }

        var yearMatch = updatedText.match(/\b(19|20)\d{2}\b/);
        if (yearMatch) {
            var yearStr = yearMatch[0];
            var yearVal = parseInt(yearStr, 10);
            var newYear = yearVal + stepInt;
            updatedText = updatedText.replace(yearStr, String(newYear));
            return updatedText;
        }

        var dot3 = updatedText.match(/^(\d+)\.(\d+)\.(\d+)$/);
        var dot2 = updatedText.match(/^(\d+)\.(\d+)$/);
        if (dot3 || dot2) {
            var a = 0,
                b = 0,
                c = null;

            if (dot3) {
                // 3要素のドット区切りは従来どおり数値として扱う
                // Three-part dot patterns stay numeric
                a = parseInt(dot3[1], 10);
                b = parseInt(dot3[2], 10);
                c = parseInt(dot3[3], 10);
                var delta3 = stepInt;
                if (SHIFT_MODE === "year") a += delta3;
                else if (SHIFT_MODE === "month") b += delta3;
                else if (SHIFT_MODE === "day" && c !== null) c += delta3;
                updatedText = a + "." + b + "." + c;
                return updatedText;
            }

            // 2要素のドット区切り（例：12.1, 12.05）
            var aStr = dot2[1];
            var bStr = dot2[2];
            var aWidth = aStr.length;
            var bWidth = bStr.length;

            a = parseInt(aStr, 10);
            b = parseInt(bStr, 10);

            // 「日付」モードでは a=月, b=日 として扱い、年は CURRENT_YEAR で仮置き
            // In "date" mode, treat a=month, b=day and use CURRENT_YEAR as the base year
            if (DOT2_MODE === "date") {
                var dateDot2 = new Date(CURRENT_YEAR, a - 1, b);

                if (SHIFT_MODE === "year") {
                    // 「年」ターゲットは「月」を増減
                    // When target is "year" (first part), shift month
                    dateDot2.setMonth(dateDot2.getMonth() + stepInt);
                } else {
                    // それ以外（「月」「日」ターゲット）は「日」を増減
                    // Other targets shift the day
                    dateDot2.setDate(dateDot2.getDate() + stepInt);
                }

                var newMonth = dateDot2.getMonth() + 1;
                var newDay = dateDot2.getDate();

                // 元の桁数を尊重してゼロ埋め（例：12.05 → 12.06）
                // Respect original digit widths for zero-padding
                var newMonthStr = (aWidth > 1) ? pad(newMonth, aWidth) : String(newMonth);
                var newDayStr = (bWidth > 1) ? pad(newDay, bWidth) : String(newDay);

                updatedText = newMonthStr + "." + newDayStr;
                return updatedText;
            }

            // 「数字」モードでは、対象が前半か後半かで挙動を変える
            // In "number" mode, adjust either the first or second part.
            var delta2 = stepInt;

            if (SHIFT_MODE === "year") {
                // 前半（12.1 の「12」）を純粋に加減算
                // Just increment/decrement the first part
                a += delta2;
            } else {
                // 後半（12.1 の「1」）を加減算しつつ、桁あふれは前半に繰り上げ／繰り下げ
                // When changing the second part, propagate overflow/underflow into the first part.
                var base = Math.pow(10, String(Math.abs(b)).length || 1); // 桁数から基数を決定 / base from digit length
                var total = a * base + b + delta2;

                // total を base 進数として a, b に分解し直す
                // Decompose total back into a (quotient) and b (remainder)
                var newA;
                if (total >= 0) {
                    newA = Math.floor(total / base);
                } else {
                    newA = -Math.ceil(-total / base);
                }
                var newB = total - newA * base;

                a = newA;
                b = newB;
            }

            updatedText = a + "." + b;
            return updatedText;
        }

        // 時刻 19:00 のようなパターンを増減
        // Increment/decrement time strings like 19:00
        var timeMatch = updatedText.match(/(\d{1,2}):(\d{2})/);
        if (timeMatch) {
            var hour = parseInt(timeMatch[1], 10);
            var minute = parseInt(timeMatch[2], 10);
            var deltaTime = stepInt;

            // 年モードを「時」、それ以外（月・日）を「分」として扱う
            // Treat SHIFT_MODE 'year' as hours, others as minutes
            if (SHIFT_MODE === "year") {
                hour += deltaTime;
            } else {
                minute += deltaTime;
            }

            // 分のオーバーフロー／アンダーフローを補正
            while (minute < 0) {
                minute += 60;
                hour -= 1;
            }
            while (minute >= 60) {
                minute -= 60;
                hour += 1;
            }

            // 時は 24 時間制でループさせる
            hour = ((hour % 24) + 24) % 24;

            var newTime =
                String(hour) + ":" +
                (minute < 10 ? "0" + String(minute) : String(minute));

            updatedText = updatedText.replace(timeMatch[0], newTime);
            return updatedText;
        }

        // ★ 数値抽出パターンを修正（こちらも 4 桁以上 OK）
        var numMatch = updatedText.match(/-?\d+(?:,\d{3})*(?:\.\d+)?/);
        if (numMatch) {
            var numberStr = numMatch[0];
            var rawStr = numberStr.replace(/,/g, "");
            var decimalMatch = rawStr.match(/\.(\d+)/);
            var decimalDigits = decimalMatch ? decimalMatch[1].length : 0;
            var value = parseFloat(rawStr);
            var step = decimalDigits > 0 ? 1 / Math.pow(10, decimalDigits) : 1;
            var result = value + step * stepVal;
            var formatted = addCommas(result.toFixed(decimalDigits));
            updatedText = updatedText.replace(numberStr, formatted);
        }
        return updatedText;

    } catch (e) {
        $.writeln("Error in updateLine: " + e.message);
        return text;
    }
}

function pad(num, len) {
    var str = String(num);
    while (str.length < len) str = "0" + str;
    return str;
}

function addCommas(str) {
    var parts = str.split(".");
    var intPart = parts[0];
    var decimalPart = parts.length > 1 ? "." + parts[1] : "";
    var isNegative = intPart.charAt(0) === "-";
    if (isNegative) intPart = intPart.substr(1);
    var withCommas = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return (isNegative ? "-" : "") + withCommas + decimalPart;
}

function changeValueByArrowKey(editText) {
    if (!editText) return;
    editText.addEventListener("keydown", function(event) {
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
                event.preventDefault();
            }
        }
        if (keyboard.altKey) {
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

showDialog();