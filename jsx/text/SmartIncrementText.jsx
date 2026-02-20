#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {
    /* =========================================
     * 連番複製スクリプト / Duplicate Text With Increment Numbers
     * 更新日: 2026-02-20
     * Version: v2.0.0
     *
     * 概要:
     * - 選択したテキストフレーム内の数字/英字を検出し、指定数だけ下方向へ複製して増分
     * - 日付/時刻は暦として正しい増減（繰り上がり/うるう年）＋曜日追従に対応
     * - 英字増分は「1文字（A〜Z/a〜z）」のみ対応（例: A1 はOK / AB1, Ver1 の英字は対象外）
     * - プレビューは常時ON。プレビューで増えたUndoをカウントし、閉じる/確定時に一括Undoで戻す
     * - ダイアログ位置を記憶して次回復元
     * ========================================= */

    // バージョン / Version
    var SCRIPT_VERSION = "v2.0.0";

    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: {
            ja: "連番複製",
            en: "Duplicate with Increment"
        },
        alertNoDoc: {
            ja: "ドキュメントが開かれていません。",
            en: "No document is open."
        },
        alertSelectOneTextFrame: {
            ja: "テキストフレームを1つだけ選択してください。",
            en: "Select exactly one text frame."
        },
        alertNoTarget: {
            ja: "増分対象が見つかりませんでした。\n英字は1文字（A〜Z）の場合のみ対応します。\n例: A1 はOK / AB1, Ver1 は英字増分対象になりません。",
            en: "No increment target was found.\nAlphabet increment supports only single-letter tokens (A–Z).\nExamples: A1 is OK / AB1, Ver1 are NOT alphabet targets."
        },
        alertNoToken: {
            ja: "選択されたテキストの中に「半角数字」または「英字」が見つかりませんでした。\n数字/英字を含むテキスト（例: 01, 2025/11/21, 19:00, A1）を選択してください。",
            en: "No digits or letters were found in the selected text.\nSelect text containing digits/letters (e.g., 01, 2025/11/21, 19:00, A1)."
        },
        labelCount: {
            ja: "複製数:",
            en: "Copies:"
        },
        labelStep: {
            ja: "増分:",
            en: "Step:"
        },
        labelInterval: {
            ja: "間隔:",
            en: "Spacing:"
        },
        labelTarget: {
            ja: "増分対象:",
            en: "Increment:"
        },
        labelStartOverride: {
            ja: "開始番号:",
            en: "Start:"
        },
        labelZeroPad: {
            ja: "ゼロ埋め",
            en: "Zero pad"
        },
        labelMergeOnOK: {
            ja: "確定時にテキストを結合",
            en: "Merge text on OK"
        },
        btnCancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        btnOK: {
            ja: "OK",
            en: "OK"
        },
        // target labels
        labelYear: { ja: "年", en: "Year" },
        labelMonth: { ja: "月", en: "Month" },
        labelDay: { ja: "日", en: "Day" },
        labelHour: { ja: "時", en: "Hour" },
        labelMinute: { ja: "分", en: "Minute" },
        prefixNum: { ja: "数字", en: "Num" },
        prefixAlpha: { ja: "英字", en: "Alpha" }
    };

    function L(key) {
        try {
            if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
            if (LABELS[key] && LABELS[key].en) return LABELS[key].en;
        } catch (_) { }
        return String(key);
    }

    function makeTargetLabel(kind, idx1based) {
        if (kind === "date_ymd") {
            // idx1based: 1..3
            if (idx1based === 1) return L("labelYear");
            if (idx1based === 2) return L("labelMonth");
            return L("labelDay");
        }
        if (kind === "time_hm") {
            if (idx1based === 1) return L("labelHour");
            return L("labelMinute");
        }
        if (kind === "alpha1") return L("prefixAlpha") + idx1based;
        return L("prefixNum") + idx1based;
    }

    /* 事前チェック / Pre-check */
    if (app.documents.length === 0) {
        alert(L("alertNoDoc"));
        return;
    }

    var sel = app.activeDocument.selection;
    if (sel.length !== 1 || sel[0].typename !== "TextFrame") {
        alert(L("alertSelectOneTextFrame"));
        return;
    }

    var originalObj = sel[0];
    var originalText = originalObj.contents;
    var __originalTextSnapshot = originalText; // 復帰用 / For restore

    // 位置保持（OK確定時にテキストが動いて見える問題の対策） / Preserve position
    function getTextFramePosition(tf) {
        try {
            if (tf && tf.position) return [tf.position[0], tf.position[1]];
        } catch (_) { }
        try {
            if (tf) return [tf.left, tf.top];
        } catch (_) { }
        return null;
    }

    function setTextFramePosition(tf, pos) {
        if (!tf || !pos) return;
        try { tf.position = [pos[0], pos[1]]; return; } catch (_) { }
        try { tf.left = pos[0]; tf.top = pos[1]; } catch (_) { }
    }

    // Undo後に参照が無効化されることがあるため、可能なら選択から再取得
    function refreshOriginalObjRef() {
        try {
            if (originalObj && originalObj.typename === "TextFrame") return;
        } catch (_) { }
        try {
            var s2 = app.activeDocument.selection;
            if (s2 && s2.length === 1 && s2[0].typename === "TextFrame") {
                originalObj = s2[0];
            }
        } catch (_) { }
    }

    // 数字/英字ラン（複数）を抽出してセグメント化 / Tokenize digits & letters
    // 例: "A1" -> tokens: ["A","1"]
    // 例: "AB1" -> tokens: ["AB","1"] （AB は英字増分対象外）
    // 例: "Ver1" -> tokens: ["Ver","1"] （Ver は英字増分対象外）
    var tokenRe = /[A-Za-z]+|\d+/g;
    var m;
    var lastIdx = 0;
    var segments = [];
    var tokensRaw = [];
    var tokenTypes = []; // "alpha1" | "num" | "other"

    while ((m = tokenRe.exec(originalText)) !== null) {
        segments.push(originalText.substring(lastIdx, m.index));
        tokensRaw.push(m[0]);
        tokenTypes.push(/^[A-Za-z]$/.test(m[0]) ? "alpha1" : (/^\d+$/.test(m[0]) ? "num" : "other"));
        lastIdx = m.index + m[0].length;
    }
    segments.push(originalText.substring(lastIdx));

    if (tokensRaw.length === 0) {
        alert(L("alertNoToken"));
        return;
    }

    // 復帰用（元のトークン配列） / Snapshot tokens
    var __baseTokensSnapshot = tokensRaw.slice();

    function rebuildText(tokenArr) {
        var s = segments[0];
        for (var i = 0; i < tokenArr.length; i++) {
            s += String(tokenArr[i]) + segments[i + 1];
        }
        return s;
    }

    // 書式判定（最低限：日付/時刻/一般） / Format detection
    var targetLabels = [];
    var targetIndices = []; // ラジオ候補の token index
    var patternType = "generic";

    // 和文年月日 / Japanese YMD (with optional weekday)
    var mYMD = originalText.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日(?:[（(［\[]?(?:日|月|火|水|木|金|土)[）)\]］]?)?$/);
    // スラッシュ区切り / Slash YMD (with optional weekday)
    var mSlash = originalText.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})(?:[ 　\t]*[（(［\[]?(?:日|月|火|水|木|金|土|Sun|Mon|Tue|Wed|Thu|Fri|Sat)[）)\]］]?)?$/);
    // ドット区切り3要素 / Dot 3 parts
    var mDot3 = originalText.match(/^(\d+)\.(\d+)\.(\d+)$/);
    // 時刻 / Time
    var mTime = originalText.match(/^(\d{1,2}):(\d{2})$/);

    if (mYMD || mSlash || mDot3) {
        patternType = "date_ymd";
        targetLabels = [makeTargetLabel("date_ymd", 1), makeTargetLabel("date_ymd", 2), makeTargetLabel("date_ymd", 3)];
        targetIndices = [0, 1, 2];
    } else if (mTime) {
        patternType = "time_hm";
        targetLabels = [makeTargetLabel("time_hm", 1), makeTargetLabel("time_hm", 2)];
        targetIndices = [0, 1];
    } else {
        var numCount = 0;
        var alphaCount = 0;
        for (var ti = 0; ti < tokensRaw.length; ti++) {
            if (tokenTypes[ti] === "alpha1") {
                alphaCount++;
                targetLabels.push(makeTargetLabel("alpha1", alphaCount));
                targetIndices.push(ti);
            } else if (tokenTypes[ti] === "num") {
                numCount++;
                targetLabels.push(makeTargetLabel("num", numCount));
                targetIndices.push(ti);
            }
            // other（AB, Ver など）は増分対象にしない
        }
    }

    // デフォルト増分対象 / Default target index
    var targetIndex = 0;
    if (patternType === "date_ymd") {
        targetIndex = 2; // Day
    } else if (patternType === "time_hm") {
        targetIndex = 1; // Minute
    } else {
        // 一般は最後の候補（数値があれば末尾数値、なければ末尾英字）
        if (targetIndices.length > 0) targetIndex = targetIndices[targetIndices.length - 1];
    }

    // 一般で増分対象がない場合（例: "AB" だけ等）
    if (patternType === "generic" && targetIndices.length === 0) {
        alert(L("alertNoTarget"));
        return;
    }

    // 対象トークン桁数 / Target length
    var targetLength = String(tokensRaw[targetIndex]).length;

    // アウトライン化した高さ（pt）を取得（計測用に複製して削除する） / Measure outlined height
    function measureOutlinedHeightPt(textFrame) {
        try {
            var tmp = textFrame.duplicate();
            try { tmp.selected = false; } catch (_) { }

            var outlined = tmp.createOutline(); // GroupItem
            var gb = outlined.geometricBounds;  // [L, T, R, B]
            var h = gb[1] - gb[3];

            try { outlined.remove(); } catch (_) { }
            return h;
        } catch (e) {
            return null;
        }
    }

    var outlinedHeightPt = measureOutlinedHeightPt(originalObj);
    if (outlinedHeightPt == null || isNaN(outlinedHeightPt) || outlinedHeightPt <= 0) {
        outlinedHeightPt = originalObj.height; // fallback
    }

    /* ================================
     * 単位ユーティリティ（text/units） / Unit util (text/units)
     * ================================ */
    var __unitMap = {
        0: "in",
        1: "mm",
        2: "pt",
        3: "pica",
        4: "cm",
        6: "px",
        7: "ft/in",
        8: "m",
        9: "yd",
        10: "ft"
    };

    function getUnitLabel(code, prefKey) {
        if (code === 5) {
            var hKeys = {
                "text/asianunits": true,
                "rulerType": true,
                "strokeUnits": true,
                "text/units": true
            };
            return hKeys[prefKey] ? "H" : "Q";
        }
        return __unitMap[code] || "pt";
    }

    function getPtFactorFromUnitCode(code) {
        switch (code) {
            case 0: return 72.0;                        // in
            case 1: return 72.0 / 25.4;                 // mm
            case 2: return 1.0;                         // pt
            case 3: return 12.0;                        // pica
            case 4: return 72.0 / 2.54;                 // cm
            case 5: return 72.0 / 25.4 * 0.25;          // Q or H
            case 6: return 1.0;                         // px
            case 7: return 72.0 * 12.0;                 // ft/in
            case 8: return 72.0 / 25.4 * 1000.0;        // m
            case 9: return 72.0 * 36.0;                 // yd
            case 10: return 72.0 * 12.0;                // ft
            default: return 1.0;
        }
    }

    function getUnitInfo(prefKey) {
        var code = 2;
        try { code = app.preferences.getIntegerPreference(prefKey); } catch (_) { code = 2; }
        return {
            code: code,
            label: getUnitLabel(code, prefKey),
            factor: getPtFactorFromUnitCode(code) // 1 unit = factor pt
        };
    }

    function ptToUnitValue(pt, unitInfo) {
        return pt / unitInfo.factor;
    }

    function unitValueToPt(v, unitInfo) {
        return v * unitInfo.factor;
    }

    var __textUnit = getUnitInfo("text/units");

    // 増分値を取得（0 は 1 扱い） / Get step value (0 => 1)
    function getStepValue() {
        var v = 1;
        try {
            v = parseInt(String(stepInput.text), 10);
        } catch (_) { v = 1; }
        if (isNaN(v)) v = 1;
        if (v === 0) v = 1;
        return v;
    }

    // プレビュー管理用 / Preview objects
    var previewObjects = [];

    /* =========================================
     * プレビュー用ヒストリー管理 / Preview history manager
     * - プレビューで増えたUndoステップ数をカウントし、閉じる/確定時に一括Undoで戻す
     * - app.undo() を優先し、失敗時は executeMenuCommand('undo') にフォールバック
     * ========================================= */
    var __suspendPreview = false;
    var __closingByOK = false;

    var PreviewHistory = (function () {
        var _count = 0;
        var _MAX_UNDO = 300; // 安全上限 / safety cap

        function _undoOnce() {
            try {
                if (app.undo) {
                    app.undo();
                    return true;
                }
            } catch (_) { }
            try {
                app.executeMenuCommand("undo");
                return true;
            } catch (e) { }
            return false;
        }

        return {
            start: function () { _count = 0; },
            bump: function () { _count++; if (_count > _MAX_UNDO) _count = _MAX_UNDO; },
            bumpBy: function (n) {
                var k = 0;
                try { k = parseInt(n, 10); } catch (_) { k = 0; }
                if (isNaN(k) || k <= 0) return;
                _count += k;
                if (_count > _MAX_UNDO) _count = _MAX_UNDO;
            },
            undo: function () {
                for (var i = 0; i < _count && i < _MAX_UNDO; i++) {
                    if (!_undoOnce()) break;
                }
                _count = 0;
            },
            getCount: function () { return _count; }
        };
    })();

    /* ダイアログ作成 / Build dialog */
    var win = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);

    /* ダイアログ位置の記憶 / Remember dialog position */
    var __WINPOS_KEY = "dupTextWithIncrementNumbers_v2_dialog_pos";

    function __readDialogPos() {
        try {
            var s = app.preferences.getStringPreference(__WINPOS_KEY);
            if (!s) return null;
            var p = s.split(",");
            if (p.length !== 2) return null;
            var x = parseFloat(p[0]);
            var y = parseFloat(p[1]);
            if (isNaN(x) || isNaN(y)) return null;
            return { x: x, y: y };
        } catch (_) { }
        return null;
    }

    function __writeDialogPos(winObj) {
        try {
            if (!winObj || !winObj.location) return;
            var x = winObj.location[0];
            var y = winObj.location[1];
            if (x == null || y == null) return;
            app.preferences.setStringPreference(__WINPOS_KEY, String(x) + "," + String(y));
        } catch (_) { }
    }

    // 前回位置を復元（失敗したらデフォルトのまま） / Restore previous position
    try {
        var __p = __readDialogPos();
        if (__p) win.location = [__p.x, __p.y];
    } catch (_) { }

    win.orientation = "column";
    win.alignChildren = "fill";

    /* 入力エリア / Inputs */
    var mainGroup = win.add("group");
    mainGroup.orientation = "column";
    mainGroup.alignChildren = "left";

    // 複製数 / Copies
    var group1 = mainGroup.add("group");
    var stCount = group1.add("statictext", undefined, L("labelCount"));
    stCount.preferredSize.width = 60;
    stCount.justify = "right";
    var countInput = group1.add("edittext", undefined, "5");
    countInput.characters = 4;

    // 増分 / Step
    var groupStep = mainGroup.add("group");
    var stStep = groupStep.add("statictext", undefined, L("labelStep"));
    stStep.preferredSize.width = 60;
    stStep.justify = "right";
    var stepInput = groupStep.add("edittext", undefined, "1");
    stepInput.characters = 4;

    // 間隔 / Spacing
    var group2 = mainGroup.add("group");
    var stInterval = group2.add("statictext", undefined, L("labelInterval"));
    stInterval.preferredSize.width = 60;
    stInterval.justify = "right";

    // デフォルト間隔 = 文字サイズの1.5倍（実際の間隔） / Default spacing = fontSize * 1.5 (actual)
    // UI表示は「実際の間隔 - 文字サイズ」 / UI shows (actual - fontSize)
    var fontSizePt = originalObj.textRange.characterAttributes.size;
    if (isNaN(fontSizePt) || fontSizePt <= 0) fontSizePt = 10;

    var defaultOffsetPt = fontSizePt * 1.5;          // actual
    var defaultGapPt = defaultOffsetPt - fontSizePt; // UI
    if (defaultGapPt < 0) defaultGapPt = 0;

    var defaultGapUnit = ptToUnitValue(defaultGapPt, __textUnit);
    var offsetInput = group2.add("edittext", undefined, defaultGapUnit.toFixed(1));
    offsetInput.characters = 4;

    var stOffsetUnit = group2.add("statictext", undefined, __textUnit.label);

    // 増分対象 / Increment target
    // 複数候補があるときだけラジオ表示 / Show radios only when multiple
    var groupTarget = mainGroup.add("group");
    groupTarget.orientation = "row";
    groupTarget.alignChildren = ["left", "center"];

    var stTarget = groupTarget.add("statictext", undefined, L("labelTarget"));
    stTarget.preferredSize.width = 60;
    stTarget.justify = "right";

    var gTargetRadios = groupTarget.add("group");
    gTargetRadios.orientation = "row";
    gTargetRadios.alignChildren = ["left", "center"];

    var rbTargets = [];
    for (var _ti = 0; _ti < targetLabels.length; _ti++) {
        var rb = gTargetRadios.add("radiobutton", undefined, targetLabels[_ti]);
        var idxToken = (targetIndices && targetIndices.length > 0) ? targetIndices[_ti] : _ti;
        rb.value = (idxToken === targetIndex);
        rbTargets.push(rb);
    }

    groupTarget.visible = (targetLabels.length > 1);

    // 開始番号 / Start override
    var groupStart = mainGroup.add("group");
    groupStart.orientation = "row";
    groupStart.alignChildren = ["left", "center"];

    var chkStartOverride = groupStart.add("checkbox", undefined, L("labelStartOverride"));
    var startInput = groupStart.add("edittext", undefined, __baseTokensSnapshot[targetIndex]);
    startInput.characters = 6;
    startInput.enabled = false;

    function onTargetIndexChanged(newIndex) {
        targetIndex = (targetIndices && targetIndices.length > 0) ? targetIndices[newIndex] : newIndex;

        var raw = __baseTokensSnapshot[targetIndex];
        targetLength = String(raw).length;

        startInput.text = String(raw);
        startInput.enabled = chkStartOverride.value;

        applyStartNumberToOriginal();
        updatePreview();
    }

    for (var _ri = 0; _ri < rbTargets.length; _ri++) {
        (function (idx) {
            rbTargets[idx].onClick = function () { onTargetIndexChanged(idx); };
        })(_ri);
    }

    // チェックONのときだけ入力可能 / Enable only when checked
    chkStartOverride.onClick = function () {
        startInput.enabled = chkStartOverride.value;
        applyStartNumberToOriginal();
        updatePreview();
    };

    // ゼロ埋め / Zero pad (default ON)
    var groupZeroPad = mainGroup.add("group");
    groupZeroPad.orientation = "row";
    groupZeroPad.alignChildren = ["left", "center"];

    var chkZeroPad = groupZeroPad.add("checkbox", undefined, L("labelZeroPad"));
    chkZeroPad.value = true;
    chkZeroPad.onClick = function () {
        applyStartNumberToOriginal();
        updatePreview();
    };

    // 確定時にテキストを結合 / Merge text on OK (default OFF)
    var groupMergeText = mainGroup.add("group");
    groupMergeText.orientation = "row";
    groupMergeText.alignChildren = ["left", "center"];

    var chkMergeTextOnOK = groupMergeText.add("checkbox", undefined, L("labelMergeOnOK"));
    chkMergeTextOnOK.value = false;

    /* ボタンエリア（OKを右寄せ） / Buttons (OK aligned right) */
    var btnGroup = win.add("group");
    btnGroup.orientation = "row";
    btnGroup.alignment = "right";

    var cancelBtn = btnGroup.add("button", undefined, L("btnCancel"), { name: "cancel" });
    var okBtn = btnGroup.add("button", undefined, L("btnOK"), { name: "ok" });

    // ↑↓ / Shift+↑↓ / Option+↑↓ で値を増減 / Change value by arrow keys
    function changeValueByArrowKey(editText, allowNegative, onChanged, forceInteger) {
        if (!editText) return;
        if (allowNegative == null) allowNegative = false;
        if (forceInteger == null) forceInteger = false;

        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                if (event.keyName === "Up") value = Math.ceil((value + 1) / delta) * delta;
                else value = Math.floor((value - 1) / delta) * delta;
            } else if (keyboard.altKey) {
                delta = 0.1;
                value += (event.keyName === "Up") ? delta : -delta;
            } else {
                value += (event.keyName === "Up") ? 1 : -1;
            }

            // rounding
            if (!forceInteger && keyboard.altKey) value = Math.round(value * 10) / 10;
            else value = Math.round(value);

            if (!allowNegative && value < 0) value = 0;

            event.preventDefault();
            editText.text = value;

            if (typeof onChanged === "function") onChanged();
        });
    }

    // ↑↓キー / Arrow key support
    changeValueByArrowKey(countInput, false, updatePreview, true);
    changeValueByArrowKey(offsetInput, false, updatePreview, false);
    changeValueByArrowKey(startInput, false, function () { applyStartNumberToOriginal(); updatePreview(); }, true);

    changeValueByArrowKey(stepInput, true, updatePreview, true);

    /* ゼロ埋め / Zero padding */
    function zeroPad(num, len) {
        var str = String(num);
        while (str.length < len) str = "0" + str;
        return str;
    }

    // 「ゼロ埋め」ON時の桁数（最大値に合わせて自動拡張） / Dynamic pad length
    function getDynamicPadLengthFor(baseStartNum, baseLen) {
        var len = baseLen;
        try {
            var dupCount = parseInt(countInput.text, 10);
            if (isNaN(dupCount) || dupCount < 0) return len;

            var stepVal = 1;
            try { stepVal = getStepValue(); } catch (_) { stepVal = 1; }

            var endVal = baseStartNum + (dupCount * stepVal);
            var dynLen = Math.max(String(baseStartNum).length, String(endVal).length);
            if (dynLen > len) len = dynLen;
        } catch (_) { }
        return len;
    }

    function formatNumberByOption(num, baseStartNum, baseLen) {
        try {
            if (chkZeroPad && chkZeroPad.value) {
                var dlen = getDynamicPadLengthFor(baseStartNum, baseLen);
                return zeroPad(num, dlen);
            }
        } catch (_) { }
        return String(num);
    }

    /* 英字（1文字のみ） / Single-letter alpha helpers */
    function isAlphaToken(s) {
        return /^[A-Za-z]$/.test(String(s));
    }

    function alphaToNumber(alphaStr) {
        var up = String(alphaStr).toUpperCase();
        if (!/^[A-Z]$/.test(up)) return null;
        return up.charCodeAt(0) - 64; // A=1..Z=26
    }

    function numberToAlpha(n, isLower) {
        var num = Math.floor(Number(n));
        if (!isFinite(num) || num <= 0) num = 1;
        // 1文字のみ：1..26 でロール
        var v = ((num - 1) % 26 + 26) % 26; // safe
        var s = String.fromCharCode(65 + v);
        return isLower ? s.toLowerCase() : s;
    }

    /* 暦として正しい増減 + 曜日追従 / Calendar-correct increment + weekday follow */
    function detectWeekdayInfo(text) {
        var mm = text.match(/([（(［\[])[ 　\t]*(日|月|火|水|木|金|土|Sun|Mon|Tue|Wed|Thu|Fri|Sat)[ 　\t]*([）)\]］])/);
        if (!mm) return { has: false };
        var token = mm[2];
        var style = (token.length === 1) ? "ja" : "en";
        return { has: true, style: style, left: mm[1], right: mm[3] };
    }

    var __weekdayInfo = detectWeekdayInfo(originalText);

    function weekdayTokenByDate(d, style) {
        var ja = ["日", "月", "火", "水", "木", "金", "土"];
        var en = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
        return (style === "en") ? en[d.getDay()] : ja[d.getDay()];
    }

    function applyWeekdayToText(text, d) {
        if (!__weekdayInfo || !__weekdayInfo.has) return text;
        var tok = weekdayTokenByDate(d, __weekdayInfo.style);
        var re = /([（(［\[])[ 　\t]*(日|月|火|水|木|金|土|Sun|Mon|Tue|Wed|Thu|Fri|Sat)[ 　\t]*([）)\]］])/;
        return text.replace(re, __weekdayInfo.left + tok + __weekdayInfo.right);
    }

    function buildBaseTokensArray() {
        var arr = __baseTokensSnapshot.slice();
        if (chkStartOverride.value) {
            var raw = String(startInput.text);
            if (tokenTypes[targetIndex] === "alpha1") {
                if (isAlphaToken(raw)) arr[targetIndex] = raw;
            } else {
                var v = parseInt(raw, 10);
                if (!isNaN(v)) arr[targetIndex] = String(v);
            }
        }
        return arr;
    }

    function parseDateFromTokens(arr) {
        if (arr.length < 3) return null;
        var y = parseInt(arr[0], 10);
        var mo = parseInt(arr[1], 10);
        var da = parseInt(arr[2], 10);
        if (isNaN(y) || isNaN(mo) || isNaN(da)) return null;

        // 2桁年などは現在世紀で補完 / Normalize short year
        var yLen = String(__baseTokensSnapshot[0]).length;
        if (yLen < 4) {
            var cy = (new Date()).getFullYear();
            var base = Math.floor(cy / 100) * 100;
            y = base + y;
        }
        return new Date(y, mo - 1, da);
    }

    function formatDateToTokens(d) {
        // 表示桁は元の桁に合わせて固定（年は元の桁で下桁を使う） / Keep original digit lengths
        var yLen = String(__baseTokensSnapshot[0]).length;
        var mLen = String(__baseTokensSnapshot[1]).length;
        var dLen = String(__baseTokensSnapshot[2]).length;

        var y = d.getFullYear();
        var mo = d.getMonth() + 1;
        var da = d.getDate();

        var ys = String(y);
        if (yLen < ys.length) ys = ys.slice(ys.length - yLen);
        if (yLen > ys.length) ys = zeroPad(ys, yLen);

        var ms = String(mo);
        var ds = String(da);
        if (mLen > 1) ms = zeroPad(ms, mLen);
        if (dLen > 1) ds = zeroPad(ds, dLen);

        return [ys, ms, ds];
    }

    function addDateByUnit(baseDate, unitIdx, step) {
        var d = new Date(baseDate.getTime());
        if (unitIdx === 0) d.setFullYear(d.getFullYear() + step);
        else if (unitIdx === 1) d.setMonth(d.getMonth() + step);
        else d.setDate(d.getDate() + step);
        return d;
    }

    function parseTimeFromTokens(arr) {
        if (arr.length < 2) return null;
        var h = parseInt(arr[0], 10);
        var mi = parseInt(arr[1], 10);
        if (isNaN(h) || isNaN(mi)) return null;
        return {
            h: h,
            m: mi,
            hLen: String(__baseTokensSnapshot[0]).length,
            mLen: String(__baseTokensSnapshot[1]).length
        };
    }

    function addTimeByUnit(baseTime, unitIdx, step) {
        var total = baseTime.h * 60 + baseTime.m;
        total += (unitIdx === 0) ? (step * 60) : step;
        total = ((total % 1440) + 1440) % 1440;
        return { h: Math.floor(total / 60), m: total % 60 };
    }

    function formatTimeToTokens(t) {
        var hLen = String(__baseTokensSnapshot[0]).length;
        var mLen = String(__baseTokensSnapshot[1]).length;
        var hs = String(t.h);
        var ms = String(t.m);
        if (hLen > 1) hs = zeroPad(hs, hLen);
        if (mLen > 1) ms = zeroPad(ms, mLen);
        return [hs, ms];
    }

    /* 元テキスト更新 / Update original text */
    function applyStartNumberToOriginal() {
        var __posKeep = getTextFramePosition(originalObj);
        var arr = __baseTokensSnapshot.slice();

        // start override
        if (chkStartOverride.value) {
            var raw = String(startInput.text);
            if (tokenTypes[targetIndex] === "alpha1") {
                if (!isAlphaToken(raw)) { setTextFramePosition(originalObj, __posKeep); return; }
                arr[targetIndex] = raw;
            } else {
                var v = parseInt(raw, 10);
                if (isNaN(v)) { setTextFramePosition(originalObj, __posKeep); return; }
                arr[targetIndex] = String(v);
            }
        }

        // date
        if (patternType === "date_ymd") {
            var baseDate = parseDateFromTokens(arr);
            if (!baseDate) {
                try { originalObj.contents = rebuildText(arr); } catch (_) { }
                setTextFramePosition(originalObj, __posKeep);
                return;
            }
            var parts = formatDateToTokens(baseDate);
            arr[0] = parts[0]; arr[1] = parts[1]; arr[2] = parts[2];
            var t = rebuildText(arr);
            t = applyWeekdayToText(t, baseDate);
            try { originalObj.contents = t; } catch (_) { }
            setTextFramePosition(originalObj, __posKeep);
            return;
        }

        // time
        if (patternType === "time_hm") {
            var baseTime = parseTimeFromTokens(arr);
            if (!baseTime) {
                try { originalObj.contents = rebuildText(arr); } catch (_) { }
                setTextFramePosition(originalObj, __posKeep);
                return;
            }
            var total = baseTime.h * 60 + baseTime.m;
            total = ((total % 1440) + 1440) % 1440;
            var h = Math.floor(total / 60);
            var mm = total % 60;
            var parts2 = formatTimeToTokens({ h: h, m: mm });
            arr[0] = parts2[0]; arr[1] = parts2[1];
            try { originalObj.contents = rebuildText(arr); } catch (_) { }
            setTextFramePosition(originalObj, __posKeep);
            return;
        }

        // alpha1 (single letter)
        if (tokenTypes[targetIndex] === "alpha1") {
            var baseA = alphaToNumber(arr[targetIndex]);
            if (baseA == null) baseA = 1;
            var isLower = (String(arr[targetIndex]) === String(arr[targetIndex]).toLowerCase());
            arr[targetIndex] = numberToAlpha(baseA, isLower);
            try { originalObj.contents = rebuildText(arr); } catch (_) { }
            setTextFramePosition(originalObj, __posKeep);
            return;
        }

        // numeric
        if (chkZeroPad && chkZeroPad.value) {
            var base = parseInt(arr[targetIndex], 10);
            if (isNaN(base)) base = 0;
            arr[targetIndex] = formatNumberByOption(base, base, targetLength);
            try { originalObj.contents = rebuildText(arr); } catch (_) { }
            setTextFramePosition(originalObj, __posKeep);
            return;
        }

        // zero pad OFF: restore unless start override
        if (chkStartOverride.value) {
            var base2 = parseInt(arr[targetIndex], 10);
            if (isNaN(base2)) base2 = 0;
            arr[targetIndex] = formatNumberByOption(base2, base2, targetLength);
            try { originalObj.contents = rebuildText(arr); } catch (_) { }
        } else {
            try { originalObj.contents = __originalTextSnapshot; } catch (_) { }
        }
        setTextFramePosition(originalObj, __posKeep);
    }

    /* 生成処理 / Generate duplicates */
    function generateNumbers() {
        var dupCount = parseInt(countInput.text, 10);
        var gapVal = parseFloat(offsetInput.text);
        if (isNaN(dupCount) || isNaN(gapVal)) return [];

        var createdItems = [];

        var stepVal = getStepValue();

        // UIの［間隔］は「実際の間隔 - 文字サイズ」→ 実際の間隔 = fontSize + gap
        var gapPt = unitValueToPt(gapVal, __textUnit);
        var offsetPt = fontSizePt + gapPt;

        for (var i = 1; i <= dupCount; i++) {
            var newObj = originalObj.duplicate();
            newObj.top = originalObj.top - (offsetPt * i);
            newObj.left = originalObj.left;

            var baseArr = buildBaseTokensArray();
            var arr = baseArr.slice();

            if (patternType === "date_ymd") {
                var baseDate = parseDateFromTokens(baseArr);
                if (baseDate) {
                    var di = addDateByUnit(baseDate, targetIndex, i * stepVal);
                    var parts = formatDateToTokens(di);
                    arr[0] = parts[0]; arr[1] = parts[1]; arr[2] = parts[2];
                    var txt = rebuildText(arr);
                    txt = applyWeekdayToText(txt, di);
                    newObj.contents = txt;
                } else {
                    var baseStartNum = parseInt(baseArr[targetIndex], 10);
                    if (isNaN(baseStartNum)) baseStartNum = 0;
                    arr[targetIndex] = String(baseStartNum + (i * stepVal));
                    newObj.contents = rebuildText(arr);
                }
            } else if (patternType === "time_hm") {
                var bt = parseTimeFromTokens(baseArr);
                if (bt) {
                    var ti = addTimeByUnit(bt, targetIndex, i * stepVal);
                    var parts2 = formatTimeToTokens(ti);
                    arr[0] = parts2[0]; arr[1] = parts2[1];
                    newObj.contents = rebuildText(arr);
                } else {
                    var baseStartNum2 = parseInt(baseArr[targetIndex], 10);
                    if (isNaN(baseStartNum2)) baseStartNum2 = 0;
                    arr[targetIndex] = String(baseStartNum2 + (i * stepVal));
                    newObj.contents = rebuildText(arr);
                }
            } else {
                // generic
                if (tokenTypes[targetIndex] === "alpha1") {
                    var baseAlphaStr = String(baseArr[targetIndex]);
                    var baseA = alphaToNumber(baseAlphaStr);
                    if (baseA == null) baseA = 1;

                    // 開始番号の上書き（英字・1文字のみ）
                    if (chkStartOverride.value) {
                        var rawA = String(startInput.text);
                        var tmpA = alphaToNumber(rawA);
                        if (tmpA != null) baseA = tmpA;
                    }

                    var isLower = (baseAlphaStr === baseAlphaStr.toLowerCase());
                    arr[targetIndex] = numberToAlpha(baseA + (i * stepVal), isLower);
                    newObj.contents = rebuildText(arr);
                } else {
                    var baseStartNum3 = parseInt(baseArr[targetIndex], 10);
                    if (isNaN(baseStartNum3)) baseStartNum3 = 0;
                    var currentVal3 = baseStartNum3 + (i * stepVal);
                    var currentNumStr = formatNumberByOption(currentVal3, baseStartNum3, targetLength);
                    arr[targetIndex] = currentNumStr;
                    newObj.contents = rebuildText(arr);
                }
            }

            createdItems.push(newObj);
        }
        return createdItems;
    }

    /* OK確定時：各行を改行して1つのテキストに統合 / Merge lines into original on OK */
    function mergeLinesIntoOriginalAndRemovePreviews() {
        var __posKeep = getTextFramePosition(originalObj);
        var lines = [];
        try { lines.push(String(originalObj.contents)); } catch (_) { lines.push(""); }

        for (var i = 0; i < previewObjects.length; i++) {
            try { lines.push(String(previewObjects[i].contents)); } catch (_) { lines.push(""); }
        }

        try { originalObj.contents = lines.join("\r"); } catch (_) { }

        // 行送り（leading）を設定 / Set leading
        // 行送り = 文字サイズ(pt) + ［間隔］(text/units→pt)
        try {
            var fsPt = originalObj.textRange.characterAttributes.size;
            if (isNaN(fsPt) || fsPt <= 0) fsPt = fontSizePt;

            var gapVal = parseFloat(offsetInput.text);
            if (isNaN(gapVal)) gapVal = 0;
            var gapPt = unitValueToPt(gapVal, __textUnit);

            var leadingPt = fsPt + gapPt;
            if (!isNaN(leadingPt) && leadingPt > 0) {
                try { originalObj.textRange.characterAttributes.autoLeading = false; } catch (_) { }
                try { originalObj.textRange.characterAttributes.leading = leadingPt; } catch (_) { }
            }
        } catch (_) { }

        // 統合後は複製したテキストを削除 / Remove merged duplicates
        for (var j = previewObjects.length - 1; j >= 0; j--) {
            try { previewObjects[j].remove(); } catch (_) { }
        }
        previewObjects = [];
        setTextFramePosition(originalObj, __posKeep);
    }

    /* プレビュー削除 / Clear preview */
    function clearPreview() {
        var removed = 0;
        try { removed = previewObjects ? previewObjects.length : 0; } catch (_) { removed = 0; }
        for (var i = (previewObjects ? previewObjects.length - 1 : -1); i >= 0; i--) {
            try { previewObjects[i].remove(); } catch (e) { }
        }
        previewObjects = [];
        return removed;
    }

    /* プレビュー更新（常時ON） / Update preview (always on) */
    function updatePreview() {
        if (__suspendPreview) return;

        // 元テキストを現在UI状態で更新（位置は内部で保持/復元）
        applyStartNumberToOriginal();

        // 開始番号が有効で不正値ならプレビューは出さず、既存プレビューだけ消す
        if (chkStartOverride.value) {
            var raw = String(startInput.text);
            if (tokenTypes[targetIndex] === "alpha1") {
                if (!isAlphaToken(raw)) {
                    try { clearPreview(); } catch (_) { }
                    app.redraw();
                    return;
                }
            } else {
                var t = parseInt(raw, 10);
                if (isNaN(t)) {
                    try { clearPreview(); } catch (_) { }
                    app.redraw();
                    return;
                }
            }
        }

        // プレビュー再生成（Undo を使わず remove ベースで安定化）
        try { clearPreview(); } catch (_) { }
        previewObjects = generateNumbers();
        app.redraw();
    }

    /* イベントリスナー / Event listeners */
    countInput.onChanging = updatePreview;
    offsetInput.onChanging = updatePreview;
    stepInput.onChanging = updatePreview;
    startInput.onChanging = function () { applyStartNumberToOriginal(); updatePreview(); };

    cancelBtn.onClick = function () {
        __suspendPreview = true;
        var __posKeep = getTextFramePosition(originalObj);
        try { clearPreview(); } catch (_) { }
        try { originalObj.contents = __originalTextSnapshot; } catch (_) { }
        setTextFramePosition(originalObj, __posKeep);
        __suspendPreview = false;
        try { __writeDialogPos(win); } catch (_) { }
        win.close();
    };

    okBtn.onClick = function () {
        __closingByOK = true;
        __suspendPreview = true;

        var __posKeepOK = getTextFramePosition(originalObj);

        // 確定：元テキストを現在UI状態で更新
        __suspendPreview = false;
        applyStartNumberToOriginal();
        setTextFramePosition(originalObj, __posKeepOK);

        // プレビューで見えているものをそのまま確定（既に生成済み）
        // ※ 念のためプレビューが空なら生成する
        if (!previewObjects || previewObjects.length === 0) {
            previewObjects = generateNumbers();
        }

        // 確定時にテキストを結合
        if (chkMergeTextOnOK && chkMergeTextOnOK.value) {
            mergeLinesIntoOriginalAndRemovePreviews();
        }

        setTextFramePosition(originalObj, __posKeepOK);

        // 管理対象外
        previewObjects = [];

        try { __writeDialogPos(win); } catch (_) { }
        win.close();
    };

    // タイトルバーの×で閉じた場合もプレビューを消し、元テキストを復帰 / Close (X)
    win.onClose = function () {
        try { __writeDialogPos(win); } catch (_) { }
        if (__closingByOK) return;
        __suspendPreview = true;
        var __posKeep = getTextFramePosition(originalObj);
        try { clearPreview(); } catch (_) { }
        try { originalObj.contents = __originalTextSnapshot; } catch (_) { }
        setTextFramePosition(originalObj, __posKeep);
        __suspendPreview = false;
    };

    /* ダイアログ表示 / Show dialog */
    updatePreview();
    win.show();

})();