#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの中から特定の文字だけを対象に、フォントサイズ・水平／垂直比率・
ベースライン・カーニング・トラッキング・文字揃えをまとめて調整するスクリプトです。

- 文字種プリセット（ひらがな／助詞／数字／記号など）または手入力で対象文字を指定
- 調整結果はライブプレビューで確認できる（空欄の項目は適用しない）
- 対象文字を切り替えると調整が積み重なる（重ねて指定）
- 「すべてリセット」で標準値に揃え、「リセット」で開いた直後の状態に戻す

### Overview

Adjusts font size, horizontal/vertical scale, baseline, kerning, tracking and
character alignment for only the targeted characters within the current text selection.

- Pick target characters via class presets (hiragana, particles, digits, symbols) or by typing
- Live preview; fields left blank are not applied
- Switching the target stacks adjustments on top of each other
- “Reset all” snaps everything to standard values; “Reset” returns to the just-opened state

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.4";

// =========================================
// ユーザー設定 / User settings
// =========================================
var PANEL_MARGINS = [16, 20, 16, 12];   // パネル余白 / Panel margins
var PANEL_SPACING = 8;                   // パネル内の標準間隔 / Default spacing inside panels
var LABEL_WIDTH = 118;                   // ラベル幅（揃える・全角約1文字分の余裕を追加）/ Unified label width (+~1 full-width char)
var DIALOG_OPACITY = 0.98;               // ダイアログ透明度 / Dialog opacity
var DIALOG_OFFSET_X = 0;               // 表示位置の横オフセット / Horizontal offset on show
var PREVIEW_INTERVAL_MS = 80;            // プレビュー間引き間隔 / Preview throttle interval

// =========================================
// ローカライズ / Localization
// =========================================

/* 言語判定 / Detect UI language */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

/* ラベル定義（カテゴリ別）/ Label definitions grouped by category */
var LABELS = {
    dialog: {
        title: { ja: "特定文字のサイズ・ベースライン調整", en: "Targeted Font Size & Baseline Adjuster" }
    },
    panel: {
        target: { ja: "対象文字", en: "Target Char" },
        fontSizeAdjust: { ja: "フォントサイズの調整", en: "Font Size Adjustment" },
        adjust: { ja: "調整（その他）", en: "Other Adjustments" },
        object: { ja: "調整（オブジェクト単位）", en: "Object-level Adjustments" }
    },
    target: {
        hiragana: { ja: "ひらがな", en: "Hiragana" },
        particle: { ja: "助詞", en: "Particles" },
        date: { ja: "日付（年月日）", en: "Date (Y/M/D)" },
        yenEtc: { ja: "¥、円、％", en: "¥, 円, %" },
        hyphen: { ja: "-（ハイフン）", en: "- (hyphen)" },
        comma: { ja: ",（カンマ）", en: ", (comma)" },
        colon: { ja: "：（コロン）", en: ": (colon)" },
        digits: { ja: "数字", en: "Digits" },
        alphabet: { ja: "アルファベット", en: "Alphabet" },
        alphaNum: { ja: "アルファベット＋数字", en: "Alphabet + digits" },
        specify: { ja: "指定", en: "Specify" }
    },
    field: {
        fontSize: { ja: "フォントサイズ", en: "Font Size" },
        scale: { ja: "水平比率/垂直比率", en: "Scale" },
        apparent: { ja: "見かけ", en: "Apparent" },
        baseline: { ja: "ベースライン", en: "Baseline" },
        kerning: { ja: "カーニング", en: "Kerning" },
        tracking: { ja: "トラッキング", en: "Tracking" },
        alignment: { ja: "文字揃え", en: "Alignment" },
        autoKern: { ja: "自動カーニング", en: "Auto Kerning" }
    },
    align: {
        keep: { ja: "変更しない", en: "No change" },
        romanBaseline: { ja: "欧文ベースライン", en: "Roman Baseline" },
        top: { ja: "仮想ボディの上{slash}右", en: "Embox Top{slash}Right" },
        center: { ja: "中央", en: "Embox Centre" },
        bottom: { ja: "仮想ボディの下{slash}左", en: "Embox Bottom{slash}Left" },
        icfTop: { ja: "平均字面の上{slash}右", en: "ICF Box Top{slash}Right" },
        icfBottom: { ja: "平均字面の下{slash}左", en: "ICF Box Bottom{slash}Left" }
    },
    autoKern: {
        keep: { ja: "変更しない", en: "No change" },
        mono: { ja: "和文等幅", en: "Japanese Equal Width" },
        zero: { ja: "0", en: "0" },
        metrics: { ja: "メトリクス", en: "Metrics" },
        optical: { ja: "オプティカル", en: "Optical" }
    },
    button: {
        ok: { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        resetAll: { ja: "すべてリセット", en: "Reset all" },
        resetOpened: { ja: "リセット", en: "Reset" },
        toApparent: { ja: "実サイズ↔見かけ", en: "Actual ↔ Apparent" },
        toApparentTip: { ja: "サイズ×比率を見かけサイズとして実フォントサイズに焼き込み、比率を100%にします。もう一度押すと焼き込み前の比率付き状態に戻ります", en: "Bakes size × scale into the actual font size at 100%. Press again to restore the previous scaled state." }
    },
    alert: {
        selectText: { ja: "テキストを選択してください", en: "Please select text" }
    },
    tip: {
        targetChar: { ja: "ここに含まれる文字だけが調整対象になります。プリセットを選ぶと自動入力され、手入力するときは「指定」を選びます。", en: "Only the characters listed here are adjusted. Picking a preset fills this in; choose “Specify” to type your own." },
        scale: { ja: "水平比率・垂直比率を同じ値でまとめて設定します。", en: "Sets the horizontal and vertical scale together to the same value." },
        apparent: { ja: "フォントサイズ×比率で計算した、実際の見た目のサイズです。", en: "The actual visual size, computed as font size × scale." },
        baseline: { ja: "文字のベースラインを上下に移動します（正で上、負で下）。", en: "Shifts the baseline up (positive) or down (negative)." },
        kerning: { ja: "隣り合う文字の間隔を 1/1000em 単位で調整します。", en: "Adjusts spacing between adjacent characters in 1/1000 em." },
        tracking: { ja: "選択文字全体の字間を 1/1000em 単位で調整します。", en: "Adjusts overall letter spacing in 1/1000 em." },
        alignment: { ja: "文字の縦方向の揃え基準を変更します（「変更しない」で現状維持）。文字単位ではなくテキストオブジェクト全体に適用します。", en: "Changes the vertical alignment basis (“No change” keeps the current setting). Applies to the whole text object, not per character." },
        autoKern: { ja: "テキストオブジェクト全体の自動カーニング方式（和文等幅／0／メトリクス／オプティカル）を設定します。", en: "Sets the auto-kerning method (Japanese equal width / 0 / metrics / optical) for the whole text object." },
        resetAll: { ja: "フォントサイズを選択範囲の先頭文字に合わせ、比率100%・ベースライン0・カーニング0・トラッキング0・文字揃え中央に揃えて適用します。", en: "Matches font size to the first character of the selection, sets scale to 100%, baseline/kerning/tracking to 0 and alignment to center, then applies." },
        resetOpened: { ja: "編集中の調整だけを取り消して空欄に戻します。重ねて確定した分は残ります。", en: "Discards only the adjustment being edited; committed (stacked) ones remain." }
    }
};

/* 言語に応じたラベル文字列を取得し {slash} を / に置換 / Resolve a label string and replace {slash} with / */
function getLocalizedText(entry) {
    return entry[currentLanguage].replace(/\{slash\}/g, "/");
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(entry) {
    return getLocalizedText(entry) + (currentLanguage === "ja" ? "：" : ":");
}

// =========================================
// 単位 / Units
// =========================================

/* 単位コードとプリファレンスキーに応じて単位ラベルを返す / Get unit label for code and pref key */
function getUnitLabel(code, prefKey) {
    var unitMap = {
        0: "in",
        1: "mm",
        2: "pt",
        3: "pica",
        4: "cm",
        5: "Q/H",
        6: "px",
        7: "ft/in",
        8: "m",
        9: "yd",
        10: "ft"
    };
    if (code === 5) {
        var hKeys = {
            "text/asianunits": true,
            "rulerType": true,
            "strokeUnits": true
        };
        return hKeys[prefKey] ? "H" : "Q";
    }
    return unitMap[code] || "pt";
}

/* 単位コード1つ分のポイント数を返す（pt換算係数）/ Points per one unit (pt conversion factor) */
function getPointsPerUnit(code) {
    var MM = 72 / 25.4; // 1mm = 72/25.4 pt
    var ptPerUnit = {
        0: 72,        // in
        1: MM,        // mm
        2: 1,         // pt
        3: 12,        // pica
        4: MM * 10,   // cm
        5: MM * 0.25, // Q/H（0.25mm）
        6: 1,         // px（72ppi）
        7: 72,        // ft/in
        8: MM * 1000, // m
        9: 72 * 36,   // yd
        10: 72 * 12   // ft
    };
    return ptPerUnit[code] || 1;
}

// =========================================
// レイアウト補助 / Layout helpers
// =========================================

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* グループの共通設定（row/column で整列を切り替え）/ Apply shared group layout (alignChildren switches by orientation) */
function setupGroup(group, orientation, spacing) {
    var groupOrientation = orientation || "column";
    group.orientation = groupOrientation;
    /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
    group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
    group.alignment = "fill";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* ラベル＋入力欄＋単位の行を追加 / Add a "label + edittext + unit" row */
function addFieldRow(parent, labelEntry, defaultValue, unitText) {
    var row = parent.add("group");
    setupGroup(row, "row");
    var label = row.add("statictext", undefined, labelText(labelEntry));
    label.justify = "right";
    var input = row.add("edittext", undefined, defaultValue);
    input.characters = 4;
    input.justify = "right";
    var unit = row.add("statictext", undefined, unitText);
    return { row: row, label: label, input: input, unit: unit };
}

/* ラベル＋表示専用テキスト＋単位の行を追加 / Add a read-only "label + value + unit" row */
function addReadoutRow(parent, labelEntry, unitText) {
    var row = parent.add("group");
    setupGroup(row, "row");
    var label = row.add("statictext", undefined, labelText(labelEntry));
    label.justify = "right";
    var value = row.add("statictext", undefined, "--");
    value.characters = 5;
    var unit = row.add("statictext", undefined, unitText);
    return { row: row, label: label, value: value, unit: unit };
}

/* ラベル＋ドロップダウンの行を追加（optionList は {label} の配列）/ Add a "label + dropdown" row (optionList is an array of {label}) */
function addDropdownRow(parent, labelEntry, optionList, tip) {
    var row = parent.add("group");
    setupGroup(row, "row");
    var label = row.add("statictext", undefined, labelText(labelEntry));
    label.justify = "right";
    var names = [];
    for (var i = 0; i < optionList.length; i++) {
        names.push(getLocalizedText(optionList[i].label));
    }
    var dropdown = row.add("dropdownlist", undefined, names);
    dropdown.selection = 0;
    if (tip) {
        label.helpTip = tip;
        dropdown.helpTip = tip;
    }
    return { row: row, label: label, dropdown: dropdown };
}

/* 行（ラベル＋入力／表示＋単位）にまとめて helpTip を設定 / Set helpTip across a field/readout row */
function setFieldTip(field, tip) {
    field.label.helpTip = tip;
    field.unit.helpTip = tip;
    if (field.input) field.input.helpTip = tip;
    if (field.value) field.value.helpTip = tip;
}

/* 複数ラベルの幅を揃える / Set a uniform width for the given labels */
function alignLabelWidths(width, labels) {
    for (var i = 0; i < labels.length; i++) {
        labels[i].preferredSize.width = width;
    }
}

/* 上下キーによる数値の増減を有効にする / Enable value increment/decrement by arrow keys */
function changeValueByArrowKey(editText, options) {
    editText.addEventListener("keydown", function (event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;
        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;
        var precision = 0;

        if (keyboard.shiftKey) {
            delta = options.shiftStep || 10;
        } else if (keyboard.altKey) {
            delta = options.altStep || 0.1;
            precision = 1;
        } else {
            delta = options.step || 1;
        }

        if (event.keyName == "Up") {
            if (keyboard.shiftKey) {
                value = Math.floor(value / delta) * delta + delta;
            } else {
                value += delta;
            }
        } else if (event.keyName == "Down") {
            if (keyboard.shiftKey) {
                value = Math.ceil(value / delta) * delta - delta;
            } else {
                value -= delta;
            }
        } else {
            return;
        }

        if (precision > 0) {
            var factor = Math.pow(10, precision);
            value = Math.round(value * factor) / factor;
        } else {
            value = Math.round(value);
        }

        event.preventDefault();
        editText.text = value;
        editText.notify("onChange");
    });
}

// =========================================
// 文字判定・選択取得 / Character matching & selection
// =========================================

/* 文字が対象に含まれるか判定 / Whether a character is a target */
function isTargetChar(ch, targetChar, hasTargetChar) {
    return !hasTargetChar || targetChar.indexOf(ch.contents) !== -1;
}

/* 選択中のテキスト範囲を取得 / Get selected text ranges from current document */
function getTextSelection() {
    var doc = app.activeDocument;
    var selection = doc.selection;
    var ranges = [];
    if (!selection) return ranges;
    /* テキスト編集モードでは selection が配列でなく TextRange になる / In text-edit mode the selection is a TextRange, not an array */
    if (selection.constructor.name === "TextRange") {
        ranges.push(selection);
        return ranges;
    }
    if (selection.length === 0) return ranges;
    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        if (item.constructor.name === "TextFrame") {
            ranges.push(item.textRange);
        } else if (item.constructor.name === "TextRange") {
            ranges.push(item);
        }
    }
    return ranges;
}

/* 英数字を除いたユニークな文字列を取得 / Get unique non-alphanumeric characters from text */
function getUniqueNonAlphanumerics(text) {
    var stripped = text.replace(/[0-9A-Za-z]/g, "");
    var result = "";
    for (var i = 0; i < stripped.length; i++) {
        var ch = stripped.charAt(i);
        if (result.indexOf(ch) === -1) {
            result += ch;
        }
    }
    return result;
}

/* テキストからユニークなひらがなを取得 / Get unique hiragana characters from text */
function getUniqueHiragana(text) {
    var result = "";
    for (var i = 0; i < text.length; i++) {
        var code = text.charCodeAt(i);
        if (code >= 0x3041 && code <= 0x3096) {
            var ch = text.charAt(i);
            if (result.indexOf(ch) === -1) {
                result += ch;
            }
        }
    }
    return result;
}

/* 複数文字列に共通する文字を取得 / Get characters common to all strings */
function getCommonCharacters(textArray) {
    if (textArray.length === 0) return "";
    var common = textArray[0];
    for (var i = 1; i < textArray.length; i++) {
        var current = textArray[i];
        var temp = "";
        for (var j = 0; j < common.length; j++) {
            var ch = common.charAt(j);
            if (current.indexOf(ch) !== -1 && temp.indexOf(ch) === -1) {
                temp += ch;
            }
        }
        common = temp;
        if (common.length === 0) break;
    }
    return common;
}

/* 対象文字に一致する文字だけにコールバックを適用 / Apply callback to matching characters in a range */
function applyToTargetCharacters(range, targetChar, hasTargetChar, callback) {
    var chars = range.characters;
    for (var j = 0; j < chars.length; j++) {
        var ch = chars[j];
        if (isTargetChar(ch, targetChar, hasTargetChar)) {
            callback(ch);
        }
    }
}

// =========================================
// エラー処理 / Error reporting
// =========================================

/* 直近に表示したエラーメッセージ（連続する同一エラーの抑制用）/ Last shown error message (to suppress repeats) */
var lastReportedError = "";

/* 同一メッセージの連続表示を抑制してアラート。プレビューの連打でアラートが溢れるのを防ぐ
   Alert while suppressing repeats of the same message, so throttled preview updates don't spam alerts */
function reportError(prefix, e) {
    var message = prefix + String(e);
    if (message === lastReportedError) return;
    lastReportedError = message;
    alert(message);
}

/* 正常に処理できたらエラー抑制状態をリセット / Reset suppression once an operation succeeds */
function clearReportedError() {
    lastReportedError = "";
}

// =========================================
// PreviewManager
// プレビュー時にUndo履歴を汚さないための小さな管理クラス。
// - addStep(): 変更処理を実行してundoDepthをカウント
// - rollback(): 編集中（未確定）のプレビューだけ取り消し
// - rollbackAll(): 開いてから適用した分を確定済みも含めてすべて取り消し
// - confirm(finalAction): OK時に編集中のプレビューだけ一度戻し、本適用を1回実行して確定する
// 「重ねて指定」した確定済みレイヤー（commitCurrentLayer で undoDepth=0 にした分）は戻さないため、
// それぞれ個別のUndo履歴として残る。confirm がまとめるのは編集中レイヤーの1回分だけ。
// undoDepth   = 編集中レイヤーのステップ数 / steps of the layer being edited
// appliedTotal = 開いてから適用しまだ取り消していない総ステップ数 / total applied (incl. committed) since open
// =========================================
function PreviewManager() {
    this.undoDepth = 0;
    this.appliedTotal = 0;

    this.addStep = function (func) {
        try {
            func();
            clearReportedError();
            this.undoDepth++;
            this.appliedTotal++;
            app.redraw();
        } catch (e) {
            reportError("プレビュー更新エラー / Preview update error: ", e);
        }
    };

    this.rollback = function () {
        while (this.undoDepth > 0) {
            app.undo();
            this.undoDepth--;
            this.appliedTotal--;
        }
        app.redraw();
    };

    /* 確定済みレイヤーも含めて開いた時点まで全戻し / Undo every applied step back to the just-opened state */
    this.rollbackAll = function () {
        while (this.appliedTotal > 0) {
            app.undo();
            this.appliedTotal--;
        }
        this.undoDepth = 0;
        app.redraw();
    };

    this.confirm = function (finalAction) {
        if (typeof finalAction === "function") {
            this.rollback();
            finalAction();
        } else {
            this.undoDepth = 0;
        }
    };
}

// UI入力欄の参照をまとめるオブジェクト / References to UI input controls
var uiElements = {
    sizeInput: null,
    hScaleInput: null,
    baselineInput: null,
    kerningInput: null,
    trackingInput: null,
    apparentSizeText: null
};

// =========================================
// メイン処理 / Main
// =========================================
function main() {
    if (app.documents.length <= 0) {
        return;
    }

    var targetRanges = getTextSelection();
    if (targetRanges.length === 0) {
        alert(getLocalizedText(LABELS.alert.selectText));
        return;
    }

    var allSelText = "";
    var textArray = [];
    for (var i = 0; i < targetRanges.length; i++) {
        allSelText += targetRanges[i].contents;
        textArray.push(targetRanges[i].contents);
    }

    var commonChars = getCommonCharacters(textArray);
    var defaultTargetChars = getUniqueNonAlphanumerics(commonChars);

    var previewManager = new PreviewManager();
    var unitCode = app.preferences.getIntegerPreference("text/units");
    var unitLabel = getUnitLabel(unitCode, "text/units");
    var unitFactor = getPointsPerUnit(unitCode);

    /* pt ⇄ ルーラー単位の換算 / Convert between points and ruler units */
    function ptToUnit(pt) { return pt / unitFactor; }
    function unitToPt(value) { return value * unitFactor; }

    /* 文字揃えの選択肢（value=null は「変更しない」）/ Character alignment options (value=null means "no change") */
    var alignList = [
        { label: LABELS.align.keep, value: null },
        { label: LABELS.align.romanBaseline, value: StyleRunAlignmentType.ROMANBASELINE },
        { label: LABELS.align.top, value: StyleRunAlignmentType.top },
        { label: LABELS.align.center, value: StyleRunAlignmentType.center },
        { label: LABELS.align.bottom, value: StyleRunAlignmentType.bottom },
        { label: LABELS.align.icfTop, value: StyleRunAlignmentType.icfTop },
        { label: LABELS.align.icfBottom, value: StyleRunAlignmentType.icfBottom }
    ];

    /* 自動カーニングの選択肢 / Auto-kerning options
       value=null は「変更しない」。和文等幅(kind="wabun")は確定時アクション再生、
       和文等幅は欧文のみメトリクス＝和文は等幅（METRICSROMANONLY）
       value=null = no change; Japanese equal width = metrics for Roman only (METRICSROMANONLY) */
    var autoKernList = [
        { label: LABELS.autoKern.keep, value: null },
        { label: LABELS.autoKern.mono, value: AutoKernType.METRICSROMANONLY },
        { label: LABELS.autoKern.zero, value: AutoKernType.NOAUTOKERN },
        { label: LABELS.autoKern.metrics, value: AutoKernType.AUTO },
        { label: LABELS.autoKern.optical, value: AutoKernType.OPTICAL }
    ];

    // ---- 対象文字の状態・検索 / Target-character state & lookup ----

    function getTargetCharState() {
        var targetChar = targetCharInput.text;
        var hasTargetChar = targetChar && targetChar.length > 0;
        return { targetChar: targetChar, hasTargetChar: hasTargetChar };
    }

    function findFirstMatchingChar(targetChar, hasTargetChar) {
        for (var i = 0; i < targetRanges.length; i++) {
            var chars = targetRanges[i].characters;
            for (var j = 0; j < chars.length; j++) {
                if (isTargetChar(chars[j], targetChar, hasTargetChar)) return chars[j];
            }
        }
        return null;
    }

    function calculateApparentSize(size, scale) {
        return Math.round(size * scale) / 100;
    }

    // ---- 値の取得・適用 / Read & apply values ----

    /* 対象文字すべてにコールバックを適用 / Apply a callback to every target character */
    function forEachTargetChar(state, action) {
        for (var i = 0; i < targetRanges.length; i++) {
            applyToTargetCharacters(targetRanges[i], state.targetChar, state.hasTargetChar, action);
        }
    }

    /* 現在UIの値をまとめて取得（空欄/NaNはnull）/ Collect current UI values (null when blank/NaN) */
    function getCurrentParams() {
        var sizeVal = parseFloat(uiElements.sizeInput.text);
        var scaleVal = parseFloat(uiElements.hScaleInput.text);
        var baselineVal = parseFloat(uiElements.baselineInput.text);
        var kerningVal = parseFloat(uiElements.kerningInput.text);
        var trackingVal = parseFloat(uiElements.trackingInput.text);
        var alignmentVal = alignDropdown.selection ? alignList[alignDropdown.selection.index].value : null;
        var autoKernVal = autoKernDropdown.selection ? autoKernList[autoKernDropdown.selection.index].value : null;
        return {
            size: isNaN(sizeVal) ? null : sizeVal,
            scale: isNaN(scaleVal) ? null : scaleVal,
            baseline: isNaN(baselineVal) ? null : baselineVal,
            kerning: isNaN(kerningVal) ? null : kerningVal,
            tracking: isNaN(trackingVal) ? null : trackingVal,
            alignment: alignmentVal,
            autoKern: autoKernVal
        };
    }

    /* 現在のUI状態を対象文字にまとめて適用 / Apply current UI state to target characters */
    function applyAllCurrentValues() {
        var state = getTargetCharState();
        var params = getCurrentParams();

        if (params.size !== null) {
            var sizeInPt = unitToPt(params.size);
            forEachTargetChar(state, function (ch) { ch.size = sizeInPt; });
        }
        if (params.scale !== null) {
            forEachTargetChar(state, function (ch) {
                ch.characterAttributes.horizontalScale = params.scale;
                ch.characterAttributes.verticalScale = params.scale;
            });
        }
        if (params.autoKern !== null) {
            // オブジェクト単位：選択範囲全体に自動カーニング方式を設定。手動カーニング/ベースラインより先に適用し、対象文字での NOAUTOKERN 上書きを優先させる
            // Object-level: set the kerning method on the whole selection; applied before manual kerning/baseline so their per-char NOAUTOKERN wins
            for (var ki = 0; ki < targetRanges.length; ki++) {
                targetRanges[ki].characterAttributes.kerningMethod = params.autoKern;
            }
        }
        if (params.baseline !== null) {
            var baselineInPt = unitToPt(params.baseline);
            forEachTargetChar(state, function (ch) {
                ch.characterAttributes.kerningMethod = AutoKernType.NOAUTOKERN;
                ch.characterAttributes.baselineShift = baselineInPt;
            });
        }
        if (params.kerning !== null) {
            forEachTargetChar(state, function (ch) {
                ch.characterAttributes.kerningMethod = AutoKernType.NOAUTOKERN;
                ch.kerning = params.kerning;
            });
        }
        if (params.tracking !== null) {
            forEachTargetChar(state, function (ch) {
                ch.characterAttributes.tracking = params.tracking;
            });
        }
        if (params.alignment !== null) {
            // オブジェクト単位：対象文字フィルタを無視し、選択範囲全体に適用 / Object-level: ignore the target filter, apply to the whole selection
            for (var ri = 0; ri < targetRanges.length; ri++) {
                targetRanges[ri].characterAttributes.alignment = params.alignment;
            }
        }
    }

    // ---- プレビュー / Preview ----

    /* プレビュー更新 / Update preview without polluting undo history */
    function updatePreview() {
        previewManager.rollback();
        previewManager.addStep(function () {
            applyAllCurrentValues();
        });
    }

    /* onChanging の連打で Undo/再適用が過剰にならないよう間引く / Throttle heavy preview updates */
    var lastPreviewTime = 0;
    function updatePreviewThrottled() {
        var now = (new Date()).getTime();
        if (now - lastPreviewTime < PREVIEW_INTERVAL_MS) return;
        lastPreviewTime = now;
        updatePreview();
    }

    // ---- 表示更新 / Display updates ----

    function updateApparentSizeDisplay() {
        var size = parseFloat(uiElements.sizeInput.text);
        var scale = parseFloat(uiElements.hScaleInput.text);
        if (isNaN(size) || isNaN(scale)) {
            uiElements.apparentSizeText.text = "--";
        } else {
            uiElements.apparentSizeText.text = calculateApparentSize(size, scale) + "";
        }

        var isDimmed = (scale === 100);
        apparentRow.label.enabled = !isDimmed;
        apparentRow.value.enabled = !isDimmed;
        apparentRow.unit.enabled = !isDimmed;
    }

    function updateInfoText() {
        // 対象文字の値を読み直すので、焼き込み前の保存状態（トグル）は破棄する
        // Reloading the target's values invalidates the saved pre-bake (toggle) state
        apparentToggleState = null;
        var state = getTargetCharState();
        var exampleChar = findFirstMatchingChar(state.targetChar, state.hasTargetChar);
        if (!exampleChar) {
            uiElements.sizeInput.text = "";
            uiElements.hScaleInput.text = "";
            uiElements.apparentSizeText.text = "--";
            return;
        }
        var sizeInUnit = ptToUnit(exampleChar.size);
        var horizontalScale = exampleChar.characterAttributes.horizontalScale;
        sizeInUnit = Math.round(sizeInUnit * 10) / 10;
        horizontalScale = Math.round(horizontalScale * 10) / 10;
        uiElements.sizeInput.text = sizeInUnit + "";
        uiElements.hScaleInput.text = horizontalScale + "";
        updateApparentSizeDisplay();
    }

    // ---- 値のリセット・確定 / Reset & commit values ----

    /* 調整欄を開いた直後と同じ空欄状態に戻す（サイズ・比率は updateInfoText が実値を再表示）/ Clear adjustment fields to the just-opened blank state */
    function resetAdjustValues() {
        uiElements.baselineInput.text = "";
        uiElements.kerningInput.text = "";
        uiElements.trackingInput.text = "";
        alignDropdown.selection = 0;
        autoKernDropdown.selection = 0;
    }

    /* 編集中のプレビューを確定（凍結）して次のレイヤーに重ねられるようにする / Freeze the current preview so the next change stacks on top */
    function commitCurrentLayer() {
        previewManager.undoDepth = 0;
    }

    function setTargetCharAndDim(text) {
        commitCurrentLayer();
        targetCharInput.text = text;
        targetCharInput.enabled = false;
        targetCharInput.notify("onChange");
    }

    /* リセット（開いたとき）：編集中の調整だけ取り消して空欄へ。確定済みの重ね分は残す / Reset (opened): drop only the layer being edited; keep committed layers */
    function resetToOpened() {
        previewManager.rollback();
        deselectOtherRadios(specifyRadio);
        specifyRadio.value = true;
        targetCharInput.text = defaultTargetChars;
        targetCharInput.enabled = true;
        resetAdjustValues();
        updateInfoText();   // サイズ・比率に実値を再表示（適用はしない）/ Re-display actual size/scale (without applying)
    }

    /* すべてリセット：フォントサイズを選択範囲の先頭文字に揃え、他の調整項目も標準値（比率100%/ベースライン0/カーニング0/トラッキング0/中央）にして適用。対象文字は変更しない
       Reset all: match font size to the first character of the selection and set every other adjustment to its standard value, then apply (target chars are left unchanged) */
    function resetAll() {
        apparentToggleState = null;
        // 対象文字フィルタに依存せず、選択範囲の先頭文字を基準にする / Use the first char of the selection, regardless of the target filter
        var firstChar = findFirstMatchingChar("", false);
        if (firstChar) {
            uiElements.sizeInput.text = Math.round(ptToUnit(firstChar.size) * 10) / 10;
        }
        uiElements.hScaleInput.text = "100";
        uiElements.baselineInput.text = "0";
        uiElements.kerningInput.text = "0";
        uiElements.trackingInput.text = "0";
        for (var ai = 0; ai < alignList.length; ai++) {
            if (alignList[ai].value === StyleRunAlignmentType.center) { alignDropdown.selection = ai; break; }
        }
        updatePreview();
        updateApparentSizeDisplay();
    }

    // =========================================
    // UI構築 / Build UI
    // =========================================
    var dialog = new Window("dialog", getLocalizedText(LABELS.dialog.title) + " " + SCRIPT_VERSION);
    dialog.alignChildren = "left";
    dialog.opacity = DIALOG_OPACITY;

    var contentGroup = dialog.add("group");
    contentGroup.orientation = "row";
    contentGroup.alignChildren = "top";
    contentGroup.spacing = 20;

    var leftColumn = contentGroup.add("group");
    leftColumn.orientation = "column";
    leftColumn.alignChildren = "fill";
    leftColumn.spacing = 20;

    var rightColumn = contentGroup.add("group");
    rightColumn.orientation = "column";
    rightColumn.alignChildren = "fill";
    rightColumn.spacing = 20;

    // 対象文字パネル / Target-character panel
    var targetPanel = leftColumn.add("panel", undefined, labelText(LABELS.panel.target));
    setupPanel(targetPanel);

    var targetPresetGroup = targetPanel.add("group");
    setupGroup(targetPresetGroup, "column");

    /* 文字種プリセット（chars は文字列、または評価して文字列を返す関数）/ Character-class presets */
    var targetPresets = [
        { entry: LABELS.target.hiragana, chars: function () { return getUniqueHiragana(allSelText); } },
        { entry: LABELS.target.particle, chars: "やでよりからとへにをがのは" },
        { entry: LABELS.target.date, chars: "0123456789年月日" },
        { entry: LABELS.target.yenEtc, chars: "¥￥円％%" },
        { entry: LABELS.target.hyphen, chars: "-" },
        { entry: LABELS.target.comma, chars: "," },
        { entry: LABELS.target.colon, chars: ":：" },
        { entry: LABELS.target.digits, chars: "0123456789" },
        { entry: LABELS.target.alphabet, chars: "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" },
        { entry: LABELS.target.alphaNum, chars: "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789" }
    ];

    // ラジオボタンの排他制御（グループをまたぐため手動で管理）/ Manual exclusive selection across rows
    var targetRadios = [];
    function deselectOtherRadios(selected) {
        for (var i = 0; i < targetRadios.length; i++) {
            if (targetRadios[i] !== selected) {
                targetRadios[i].value = false;
            }
        }
    }

    function addTargetRadio(rowGroup, preset) {
        var radio = rowGroup.add("radiobutton", undefined, getLocalizedText(preset.entry));
        radio.onClick = function () {
            deselectOtherRadios(radio);
            setTargetCharAndDim((typeof preset.chars === "function") ? preset.chars() : preset.chars);
        };
        targetRadios.push(radio);
        return radio;
    }

    // 1列に並べる（1行1つ）/ Lay out presets one per row
    for (var p = 0; p < targetPresets.length; p++) {
        addTargetRadio(targetPresetGroup, targetPresets[p]);
    }

    var specifyRadio = targetPresetGroup.add("radiobutton", undefined, getLocalizedText(LABELS.target.specify));
    targetRadios.push(specifyRadio);
    specifyRadio.onClick = function () {
        deselectOtherRadios(specifyRadio);
        commitCurrentLayer();
        targetCharInput.enabled = true;
        targetCharInput.notify("onChange");
    };

    var targetCharGroup = targetPanel.add("group");
    setupGroup(targetCharGroup, "row");
    var targetCharInput = targetCharGroup.add("edittext", undefined, defaultTargetChars);
    targetCharInput.characters = 16;
    targetCharInput.helpTip = getLocalizedText(LABELS.tip.targetChar);
    specifyRadio.value = true;

    // フォントサイズの調整パネル / Font-size adjustment panel
    var fontSizePanel = rightColumn.add("panel", undefined, getLocalizedText(LABELS.panel.fontSizeAdjust));
    setupPanel(fontSizePanel, 6);

    var sizeField = addFieldRow(fontSizePanel, LABELS.field.fontSize, "0", unitLabel);
    uiElements.sizeInput = sizeField.input;

    var scaleField = addFieldRow(fontSizePanel, LABELS.field.scale, "100", "%");
    uiElements.hScaleInput = scaleField.input;
    setFieldTip(scaleField, getLocalizedText(LABELS.tip.scale));

    var apparentRow = addReadoutRow(fontSizePanel, LABELS.field.apparent, unitLabel);
    uiElements.apparentSizeText = apparentRow.value;
    setFieldTip(apparentRow, getLocalizedText(LABELS.tip.apparent));

    // 焼き込み前の状態（順方向で保存→逆方向で復元）。手動でサイズ/比率を変えたら無効化
    // Pre-bake state (saved on forward, restored on back); cleared when size/scale is edited by hand
    var apparentToggleState = null;

    // 実サイズ↔見かけのトグルボタン / Toggle between actual size and apparent (baked) size
    // 順方向：サイズ×比率を実サイズに焼き込み比率100%へ。逆方向：直前の比率付き状態へ戻す
    // Forward: bake size × scale into the actual size at 100%. Back: restore the previous scaled state.
    var convertButton = fontSizePanel.add("button", undefined, getLocalizedText(LABELS.button.toApparent));
    convertButton.helpTip = getLocalizedText(LABELS.button.toApparentTip);
    convertButton.alignment = "right";
    convertButton.preferredSize.width = 150;
    convertButton.onClick = function () {
        var size = parseFloat(uiElements.sizeInput.text);
        var scale = parseFloat(uiElements.hScaleInput.text);
        if (isNaN(size) || isNaN(scale)) return;
        if (apparentToggleState !== null) {
            // 見かけ→実サイズ：焼き込み前の比率付き状態に戻す / Restore the pre-bake scaled state
            uiElements.sizeInput.text = apparentToggleState.size + "";
            uiElements.hScaleInput.text = apparentToggleState.scale + "";
            apparentToggleState = null;
        } else {
            // 実サイズ→見かけ：比率をサイズへ焼き込み100%に / Bake scale into size and reset to 100%
            apparentToggleState = { size: size, scale: scale };
            uiElements.sizeInput.text = calculateApparentSize(size, scale) + "";
            uiElements.hScaleInput.text = "100";
        }
        updatePreview();
        updateApparentSizeDisplay();
    };

    // 調整（その他）パネル / Adjust (other) panel
    var adjustPanel = rightColumn.add("panel", undefined, getLocalizedText(LABELS.panel.adjust));
    setupPanel(adjustPanel, 6);

    /* 空欄＝適用しない。初期値は空にしておく / Empty = not applied; start blank */
    var baselineField = addFieldRow(adjustPanel, LABELS.field.baseline, "", unitLabel);
    uiElements.baselineInput = baselineField.input;
    setFieldTip(baselineField, getLocalizedText(LABELS.tip.baseline));

    var kerningField = addFieldRow(adjustPanel, LABELS.field.kerning, "", "/1000");
    uiElements.kerningInput = kerningField.input;
    setFieldTip(kerningField, getLocalizedText(LABELS.tip.kerning));

    var trackingField = addFieldRow(adjustPanel, LABELS.field.tracking, "", "/1000");
    uiElements.trackingInput = trackingField.input;
    setFieldTip(trackingField, getLocalizedText(LABELS.tip.tracking));

    // 調整（オブジェクト単位）パネル — 文字単位でなくテキストオブジェクト全体に適用 / Object-level panel: applies to the whole text object, not per character
    var objectPanel = rightColumn.add("panel", undefined, getLocalizedText(LABELS.panel.object));
    setupPanel(objectPanel, 6);

    // 文字揃え（ドロップダウン）/ Character alignment (dropdown)
    var alignRow = addDropdownRow(objectPanel, LABELS.field.alignment, alignList, getLocalizedText(LABELS.tip.alignment));
    var alignLabel = alignRow.label;
    var alignDropdown = alignRow.dropdown;
    alignDropdown.onChange = function () { updatePreview(); };

    // 自動カーニング（ドロップダウン・テキストオブジェクト全体に適用）/ Auto kerning (dropdown; applies to the whole text object)
    var autoKernRow = addDropdownRow(objectPanel, LABELS.field.autoKern, autoKernList, getLocalizedText(LABELS.tip.autoKern));
    var autoKernLabel = autoKernRow.label;
    var autoKernDropdown = autoKernRow.dropdown;
    autoKernDropdown.onChange = function () { updatePreview(); };

    alignLabelWidths(LABEL_WIDTH, [
        sizeField.label, scaleField.label, apparentRow.label,
        baselineField.label, kerningField.label, trackingField.label,
        alignLabel, autoKernLabel
    ]);

    // 入力欄のイベント / Wire input events
    // サイズ・比率は「入力値をそのまま適用」。updateInfoText() で入力欄を読み直すと
    // 入力値が丸めで戻る恐れがあるため呼ばない（見かけ表示だけ更新する）
    // Apply the typed value as-is; do NOT call updateInfoText() here (re-reading the field
    // could snap the typed value back via rounding). Only refresh the apparent readout.
    uiElements.sizeInput.onChange = function () {
        apparentToggleState = null;   // 手動編集でトグル復元を無効化 / manual edit invalidates the toggle
        updatePreview();
        updateApparentSizeDisplay();
    };
    uiElements.sizeInput.onChanging = function () { updateApparentSizeDisplay(); };
    changeValueByArrowKey(uiElements.sizeInput, { step: 1, shiftStep: 10, altStep: 0.1 });

    uiElements.hScaleInput.onChange = function () {
        apparentToggleState = null;   // 手動編集でトグル復元を無効化 / manual edit invalidates the toggle
        updatePreview();
        updateApparentSizeDisplay();
    };
    uiElements.hScaleInput.onChanging = function () { updateApparentSizeDisplay(); };
    changeValueByArrowKey(uiElements.hScaleInput, { step: 1, shiftStep: 10, altStep: 5 });

    uiElements.baselineInput.onChange = updatePreview;
    uiElements.baselineInput.onChanging = updatePreviewThrottled;
    changeValueByArrowKey(uiElements.baselineInput, { step: 1, shiftStep: 10, altStep: 0.1 });

    uiElements.kerningInput.onChange = updatePreview;
    uiElements.kerningInput.onChanging = updatePreviewThrottled;
    changeValueByArrowKey(uiElements.kerningInput, { step: 1, shiftStep: 10, altStep: 0.1 });

    uiElements.trackingInput.onChange = updatePreview;
    uiElements.trackingInput.onChanging = updatePreviewThrottled;
    changeValueByArrowKey(uiElements.trackingInput, { step: 1, shiftStep: 10, altStep: 0.1 });

    targetCharInput.onChange = function () {
        updateInfoText();
        updatePreview();
    };

    // ボタン（下部・3カラム: 左=リセット / 中央=スペーサー / 右=OK・キャンセル）/ Buttons (bottom, 3 columns)
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "fill";
    buttonGroup.alignChildren = ["fill", "center"];

    var buttonLeft = buttonGroup.add("group");
    buttonLeft.alignment = ["left", "center"];
    buttonLeft.spacing = 6;
    var resetOpenedButton = buttonLeft.add("button", undefined, getLocalizedText(LABELS.button.resetOpened));
    resetOpenedButton.helpTip = getLocalizedText(LABELS.tip.resetOpened);
    var resetAllButton = buttonLeft.add("button", undefined, getLocalizedText(LABELS.button.resetAll));
    resetAllButton.helpTip = getLocalizedText(LABELS.tip.resetAll);

    var buttonCenter = buttonGroup.add("group");
    buttonCenter.alignment = ["fill", "center"];

    var buttonRight = buttonGroup.add("group");
    buttonRight.alignment = ["right", "center"];
    var cancelButton = buttonRight.add("button", undefined, getLocalizedText(LABELS.button.cancel), { name: "cancel" });
    var okButton = buttonRight.add("button", undefined, getLocalizedText(LABELS.button.ok), { name: "ok" });

    okButton.preferredSize.width = 90;
    cancelButton.preferredSize.width = 90;

    okButton.onClick = function () {
        // 編集中のプレビューを1回の適用として確定（重ねて指定した分は個別のUndo履歴として残る）
        // Commit the in-progress preview as a single apply (stacked layers remain as separate undo entries)
        previewManager.confirm(function () {
            applyAllCurrentValues();
        });
        dialog.close();
    };
    cancelButton.onClick = function () {
        // 開いてから適用した分をすべて取り消してから閉じる / Undo everything applied since open, then close
        previewManager.rollbackAll();
        dialog.close(2);
    };
    resetAllButton.onClick = function () {
        resetAll();
    };
    resetOpenedButton.onClick = function () {
        resetToOpened();
    };

    // 初期表示（updateInfoText が対象文字の実サイズを読み取って各欄を設定）/ Initial state (updateInfoText reads the actual size of the target char)
    updateInfoText();
    updateApparentSizeDisplay();

    dialog.onShow = function () {
        dialog.location = [dialog.location[0] + DIALOG_OFFSET_X, dialog.location[1]];
        uiElements.hScaleInput.active = true;
    };

    // 開いた時点では何も適用しない（現在の状態をそのまま保持）。値を変更したときだけプレビュー適用
    // Apply nothing on open (keep the current state as-is); preview only kicks in once a value changes

    dialog.show();
}

main();
